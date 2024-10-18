import pyautogui
import time

# Wait for 5 seconds
time.sleep(5)

# Get the current position of the mouse cursor
x, y = pyautogui.position()
print(f'Cursor Position after 5 seconds: X={x}, Y={y}')
