import tkinter as tk
import pyautogui
import time


# Function to simulate typing the message
def type_message(event=None):  # event is needed for key binding
    message = """
Dear Amazon Team,

According to the shipment summary, you received less of the products than the total amount we shipped.

Therefore, in line with this situation, We kindly ask you to investigate and find lost products or reimburse for the loss.

Please investigate every FNSKU missing or extra from the related shipment.

There are many units missing from the shipment.

Note: Please do not merge this case as the previous one was closed without a resolution and we no longer have the option to respond or reopen it.

Please see the attached proof of ownership/purchase and proof of delivery showing the required information. Kindly reimburse our inbound missing units under the related Shipment ID.

Thank you.
"""
    # Pause briefly before typing to allow the user to click on the input field
    time.sleep(0.2)

    # Simulate typing the message
    pyautogui.write(message, interval=0.01)


# Create the main window
root = tk.Tk()
root.title("Amazon Message Typer")

# Bind the "9" key to the type_message function
root.bind('9', type_message)

# Create a label to inform users they can press "9" to type the message
label = tk.Label(root, text='Press "9" to type the message directly', padx=10, pady=20)
label.pack()

# Run the GUI loop
root.mainloop()
