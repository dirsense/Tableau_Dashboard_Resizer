import os
import uuid
import shutil
import zipfile
import glob
import re
import pathlib
import datetime
import xml.etree.ElementTree as ET
from collections import defaultdict
from difflib import SequenceMatcher
from decimal import *

class resizer:
    def __init__(self, workbook_path: str, target_dashboard_name: str, new_dashboard_height: Decimal, new_dashboard_width: Decimal):
        self.workbook_path = workbook_path
        self.target_dashboard_name = target_dashboard_name
        self.new_dashboard_height = new_dashboard_height
        self.new_dashboard_width = new_dashboard_width
        self.workbook_name = os.path.basename(workbook_path)
        self.is_twbx = os.path.splitext(workbook_path)[1] == '.twbx'
        self.work_uuid_dir = ''
        self.twbx_workBaseDir = ''
        self.twb_path = ''
        self.final_path = ''

    def resize_process(self) -> str:
        self.create_work_dir()
        self.init_twbx()
        self.resize_dashboard_size_in_twb()
        self.set_final_path()
        self.delete_work_dir()

        return self.final_path
    
    def init_twbx(self):
        if self.is_twbx:
            self.init_twbx_work()
        else:
            self.twb_path = self.work_uuid_dir + '/' + self.workbook_name
            shutil.copy(self.workbook_path, self.twb_path)

    def resize_dashboard_size_in_twb(self):
        ET.register_namespace("user", "http://www.tableausoftware.com/xml/user")
        tree = ET.parse(self.twb_path)
        root = tree.getroot()

        BASE_SCALE = 100000

        for dash in root.iter('dashboard'):
            if dash.attrib['name'] == self.target_dashboard_name:
                dashsize = dash.find('size')

                old_dashboard_width = Decimal(dashsize.attrib['maxwidth'])
                old_dashboard_height = Decimal(dashsize.attrib['maxheight'])

                dashsize.set('maxwidth', str(self.new_dashboard_width))
                dashsize.set('minwidth', str(self.new_dashboard_width))
                dashsize.set('maxheight', str(self.new_dashboard_height))
                dashsize.set('minheight', str(self.new_dashboard_height))

                for zone in dash.iter('zone'):
                    w = Decimal(zone.attrib['w'])
                    x = Decimal(zone.attrib['x'])
                    h = Decimal(zone.attrib['h'])
                    y = Decimal(zone.attrib['y'])

                    new_w = self.round_half_up(self.round_half_up(w * (old_dashboard_width / BASE_SCALE)) / self.new_dashboard_width * BASE_SCALE)
                    new_x = self.round_half_up(self.round_half_up(x * (old_dashboard_width / BASE_SCALE)) / self.new_dashboard_width * BASE_SCALE)
                    new_h = self.round_half_up(self.round_half_up(h * (old_dashboard_height / BASE_SCALE)) / self.new_dashboard_height * BASE_SCALE)
                    new_y = self.round_half_up(self.round_half_up(y * (old_dashboard_height / BASE_SCALE)) / self.new_dashboard_height * BASE_SCALE)

                    zone.set('w', str(new_w))
                    zone.set('h', str(new_h))
                    zone.set('x', str(new_x))
                    zone.set('y', str(new_y))
                break
        
        tree.write(self.twb_path, xml_declaration=True, encoding='utf-8')

    def set_final_path(self):
        nowtime_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        self.final_path = os.path.splitext(self.workbook_name)[0] + '_resized_' + nowtime_str + ('.twbx' if self.is_twbx else '.twb')
        if self.is_twbx:
            with zipfile.ZipFile(self.final_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                for folder, subfolders, files in os.walk(self.twbx_workBaseDir):
                    for file in files:
                        if os.path.splitext(file)[1] == '.twb':
                            zf.write(self.twb_path, os.path.basename(self.twb_path))
                        else:
                            fullpath = pathlib.Path(os.path.join(folder, file))
                            relative_path = fullpath.relative_to(self.twbx_workBaseDir)
                            zf.write(fullpath, relative_path)
        else:
            shutil.copy(self.twb_path, self.final_path)

    def create_work_dir(self):
        work_dir = './work_' + str(uuid.uuid4())
        os.makedirs(work_dir, exist_ok=True)
        self.work_uuid_dir = work_dir

    def delete_work_dir(self):
        if os.path.isdir(self.work_uuid_dir):
            uuid_dirname = os.path.basename(self.work_uuid_dir)
            if (len(uuid_dirname) == 41) and (uuid_dirname[:5] == 'work_'):
                shutil.rmtree(self.work_uuid_dir)

    def round_half_up(self, t):
        return int(Decimal(t).quantize(Decimal('0'), ROUND_HALF_UP))
    
    def init_twbx_work(self):
        self.twbx_workBaseDir = self.work_uuid_dir + '/twbx_content'
        with zipfile.ZipFile(self.workbook_path) as zf:
            for info in zf.infolist():
                if not (info.flag_bits & 0x800):
                    try:
                        info.filename = info.orig_filename.encode('cp437').decode('cp932')
                    except:
                        info.filename = info.orig_filename.encode('cp437').decode('utf-8')
                    if os.sep != "/" and os.sep in info.filename:
                        info.filename = info.filename.replace(os.sep, "/")

                zf.extract(info, path=self.twbx_workBaseDir)

        twb_origin_name = os.path.splitext(os.path.basename(self.workbook_path))[0]
        twb_before_path = self.get_filepath_specify_extention(self.twbx_workBaseDir + '/*twb')
        twb_after_path = self.twbx_workBaseDir + '/' + twb_origin_name + '.twb'
        os.rename(twb_before_path, twb_after_path)
        self.twb_path = twb_after_path

        self.encode_and_rename_twbx_content('Data')
        self.encode_and_rename_twbx_content('Image')

    def get_filepath_specify_extention(self, wildcard_filepath):
        for path in glob.glob(wildcard_filepath):
            return path
        
    def encode_and_rename_twbx_content(self, target_dirname):
        target_dirpath = self.twbx_workBaseDir + '/' + target_dirname
        if os.path.isdir(target_dirpath):
            twb_paths = self.get_content_path_in_twb(target_dirname)
            twbx_dir_paths = self.get_content_path_in_twbx_dir(target_dirpath)

            for path in twbx_dir_paths:
                p = str(path)
                slash_count = p.count('/') + 1
                for i in range(slash_count):
                    if i > 0:
                        p = os.path.dirname(p)

                    testpath = os.path.basename(p)
                    if not (testpath in twb_paths[slash_count - i]):
                        try:
                            testpath_encode = testpath.encode('shift-jis').decode('utf-8', erros='ignore')
                        except:
                            testpath_encode = testpath
                        
                        twb_like_rates = {}
                        for twb_path in twb_paths[slash_count - i]:
                            twb_like_rates[twb_path] = SequenceMatcher(None, testpath_encode, twb_path).ratio()

                        if len(twb_like_rates) > 0:
                            max_rate_path = max(twb_like_rates, key=twb_like_rates.get)
                            middle_dir = os.path.dirname(p) + ('' if os.path.dirname(p) == '' else '/')
                            before = target_dirpath + '/' + middle_dir + testpath
                            after = target_dirpath + '/' + middle_dir + max_rate_path
                            os.rename(before, after)

    def get_content_path_in_twb(self, target_dirname):
        f = open(self.twb_path, 'r', encoding='utf-8')
        twb = f.read()
        f.close()

        if target_dirname == 'Image':
            pattern = '(<image-path>(Image/[^<]*)|\'(Image/[^\']*))'
        else:
            pattern = '\'(' + target_dirname + '/[^\']*)'

        d = defaultdict(set)

        for m in re.finditer(pattern, twb, re.MULTILINE):
            t = m.group(1)
            slash_count = t.count('/')
            d[slash_count].add(os.path.basename(t))
            for i in reversed(range(1, slash_count)):
                dirname = os.path.basename(os.path.dirname(t))
                d[i].add(dirname)

        return d
    
    def get_content_path_in_twbx_dir(self, target_dirpath):
        paths = []
        for folder, subfolders, files in os.walk(target_dirpath):
            for file in files:
                fullpath = pathlib.Path(os.path.join(folder, file))
                relative_path = str(fullpath.relative_to(target_dirpath)).replace(os.sep, '/')
                paths.append(relative_path)

        return paths