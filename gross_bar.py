from alive_progress import alive_bar
import time
from alive_progress.styles import showtime
from alive_progress.styles import show_bars
from alive_progress.styles import show_spinners
from alive_progress.styles import show_themes

if __name__ == '__main__':

    with alive_bar() as bar:
        for i in range(100):
            time.sleep(.1)
            bar()

    print()
    with alive_bar(100, title='Program is starting ...', bar='bubbles', length=40, spinner='radioactive') as bar:
        for i in range(100):
            time.sleep(.05)
            bar()
    print()
