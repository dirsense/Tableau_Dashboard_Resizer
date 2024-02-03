import PySimpleGUI as sg
import glob
import subprocess
from resizer import resizer
import tableaudocumentapi.xfile as tabx
import xml.etree.ElementTree as ET
import sg_utility as sgu
from decimal import *

sg.theme('DarkGrey8')

def get_current_paths_and_dashboard_sheets() -> tuple[str, str, list[str], str]:
    cur_paths = glob.glob('*.twb') + glob.glob('*.twbx')
    cur_value = '' if len(cur_paths) == 0 else cur_paths[0]

    dashlist, dashname = get_dash_elements(cur_value)

    return cur_value, cur_paths, dashlist, dashname

def get_dash_elements(twb_path: str) -> tuple[list[str], str]:
    if twb_path == '':
        return '', ''
    root: ET.Element = tabx.xml_open(twb_path).getroot()
    dashlist = []
    for dash in root.iter('dashboard'):
        dashlist.append(dash.attrib['name'])
    dashname = '' if len(dashlist) == 0 else dashlist[0]

    return dashlist, dashname

def main():
    cur_value, cur_paths, dashlist, dashname = get_current_paths_and_dashboard_sheets()

    layout = [
        [sg.Column(vertical_alignment='top', layout=[
            [sg.Text('New Width')],
            [sg.Text('New Height')],
            [sg.Text('Dashboard Name')],
            [sg.Text('TWB or TWBX Path')],
            [sg.Button('Resize', pad=((5, 5), (5, 3)))],
            [sg.Radio('Open Resize Book', key='-genop1-', group_id='genop', default=True, pad=((1, 0), (3, 0)))],
            [sg.Radio('Not Open', key='-genop2-', group_id='genop', pad=((1, 0), (0, 0)))]
        ], pad=(0, 0)),
        sg.Column(vertical_alignment='top', layout=[
            [sg.Input(key='-new_width-', size=(10, 1))],
            [sg.Input(key='-new_height-', size=(10, 1))],
            [sg.Combo(dashlist, default_value=dashname, key='-dashboard_name-', size=(50, 1))],
            [sg.Combo(cur_paths, default_value=cur_value, key='-twbpath-', size=(50, 1), enable_events=True)],
            [sg.Button('Reload Current Path', pad=((5, 5), (5, 3))),
             sg.FileBrowse('...', target='-twbpath-', file_types=(('Tableau Workbook', '*.twb*'), ),
                initial_folder='.', pad=((5, 5), (5, 3)))]
        ], pad=(0, 0))]
    ]

    window = sg.Window('Tableau Dashboard Resizer', layout)

    while True:
        event, values = window.read()

        if event is None:
            break

        elif event == '-twbpath-':
            dashlist, dashname = get_dash_elements(values['-twbpath-'])
            window['-dashboard_name-'].update(value=dashname, values=dashlist)

        elif event == 'Resize':
            p = values['-twbpath-']
            t = values['-dashboard_name-']
            h = values['-new_height-']
            w = values['-new_width-']
            if p == '' or t == '' or h == '' or w == '':
                sg.popup_error('Some fields are blank!')
            else:
                try:
                    r = resizer(workbook_path=p, target_dashboard_name=t, new_dashboard_height=Decimal(h), new_dashboard_width=Decimal(w))
                    resized_path = r.resize_process()
                except Exception as e:
                    sg.popup_error('An error occurred in the resizing process!\n\n' + e.args[0])
                else:
                    sg.popup_quick_message('Resizing process is complete!', relative_location=sgu.trailing_location(window, 0, 0), background_color='#4a6886')
                    op1 = values['-genop1-']
                    if op1:
                        subprocess.Popen(resized_path, shell=True)

        elif event == 'Reload Current Path':
            cur_value, cur_paths, dashlist, dashname = get_current_paths_and_dashboard_sheets()
            window['-twbpath-'].update(value=cur_value, values=cur_paths)
            window['-dashboard_name-'].update(value=dashname, values=dashlist)

    window.close()

if __name__ == '__main__':
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)
    except:
        pass
    main()