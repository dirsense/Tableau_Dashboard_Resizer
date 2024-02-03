import PySimpleGUI as sg

def trailing_location(window: sg.Window, relative_x: int, relative_y: int) -> tuple[int, int]:
    currloc = window.current_location()
    lastloc = window.config_last_location

    trailing_x = currloc[0] - lastloc[0] + relative_x
    trailing_y = currloc[1] - lastloc[1] + relative_y

    return trailing_x, trailing_y