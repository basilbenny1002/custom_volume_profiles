# Import necessary modules
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
import json  # For saving and loading profiles
import webbrowser  # To open URLs
import flet as ft  # For GUI components
from ctypes import cast, POINTER  # For casting COM objects
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize  # COM library for audio control

# Empty variable definitions
update_ui = None  # Placeholder for a function to update the UI
scrollable_list = ft.ListView(expand=True, spacing=10)  # Scrollable list for profiles


def save_profile(profile_name: str):
    """
    Saves audio levels of currently open processes to a JSON file with the profile name as a key.
    :param profile_name: The name of the profile to save
    """
    # Retrieve audio session details
    sessions = AudioUtilities.GetAllSessions()

    # Retrieve master volume level
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    master_volume = volume.GetMasterVolumeLevelScalar()

    # Ensure the JSON file exists
    try:
        with open("data.json", "r") as f:
            pass
    except FileNotFoundError:
        with open("data.json", "w") as f:
            pass

    # Load or initialize the JSON file
    with open("data.json", "r+") as f:
        try:
            file = json.load(f)  # Load existing profiles
        except json.JSONDecodeError:  # Handle empty or corrupted files
            file = {}
        except Exception as e:  # Catch unexpected exceptions
            print(f"Exception occurred {e}")

        # Save session volumes and master volume
        file[profile_name] = {
            session.Process.name(): session.SimpleAudioVolume.GetMasterVolume() for session in sessions
            if session.Process and session.Process.name()
        }
        file[profile_name]['master_volume'] = master_volume

    # Write the updated data to the JSON file
    print(file)
    json_string = json.dumps(file, indent=4)
    with open("data.json", "w") as f:
        f.write(json_string)


def delete_profile(profile_name):
    """
    Removes a profile from the JSON file.
    :param profile_name: The name of the profile to delete
    """
    print(profile_name)

    with open("data.json", "r+") as f:
        file = json.load(f)
        if profile_name in file:
            # Delete the specified profile
            del file[profile_name]

            # Reset the file's contents
            f.seek(0)
            f.truncate()

            # Save the updated JSON
            json.dump(file, f, indent=4)
        else:
            print("some error occurred, profile not found")

    # Update the UI to reflect changes
    update_ui()


def load_data():
    """
    Loads profiles from the JSON file and updates the scrollable list in the UI.
    """
    try:
        with open("data.json", "r") as f:
            try:
                loaded_data = json.load(f)  # Load profile data
            except json.JSONDecodeError:  # Handle corrupted files
                loaded_data = None
    except FileNotFoundError:  # Handle missing files
        loaded_data = None

    # Populate the scrollable list with profile names
    if loaded_data is not None and len(loaded_data) > 0:
        for i in loaded_data:
            scrollable_list.controls.append(
                ft.Row(
                    controls=[
                        ft.Text(i, expand=1, max_lines=1),  # Profile name
                        ft.ElevatedButton(  # Load profile button
                            text="Load Profile",
                            on_click=lambda e, profile_name=i: load_profile(profile_name),
                        ),
                        ft.IconButton(  # Delete profile button
                            icon=ft.icons.DELETE,
                            icon_color=ft.colors.RED,
                            on_click=lambda e, profile_name=i: delete_profile(profile_name),
                        ),
                    ],
                    spacing=10,
                    alignment=ft.alignment.center_left
                )
            )
            scrollable_list.controls.append(
                ft.Divider(height=1, thickness=1, color=ft.colors.GREY)  # Divider between profiles
            )
    else:
        # Message when no profiles are available
        scrollable_list.controls.append(ft.Row(controls=[ft.Text("Create a profile to get started")]))


def set_master_volume(volume_level):
    """
    Sets the master volume of the system.
    :param volume_level: The desired volume level (0.0 to 1.0)
    """
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(float(volume_level), None)


def load_profile(profile_name):
    """
    Loads the audio profile for all available processes.
    :param profile_name: The name of the profile to load
    """
    CoInitialize()  # Initialize COM library
    print(f"loaded profile for {profile_name}")

    with open("data.json", "r") as f:
        loaded_data = json.load(f)  # Load profile data
        values = loaded_data[profile_name]  # Get the selected profile
        sessions = AudioUtilities.GetAllSessions()  # Retrieve all audio sessions

        for session in sessions:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)  # Interface for audio control
            try:
                if session.Process and session.Process.name():
                    if session.Process.name() in values:
                        try:
                            volume.SetMasterVolume(float(values[session.Process.name()]), None)
                        except Exception as e:
                            print(f"Exception occurred while setting volume {e}")
            except Exception as e:
                print(f"Meh, nothing too serious {e}")

        # Set the master volume from the profile
        set_master_volume(values['master_volume'])

    CoUninitialize()  # Uninitialize COM library


def main(page: ft.Page):
    """
    Entry point for the GUI application.
    :param page: The Flet page object
    """
    global update_ui

    # Configure the window
    page.title = "Custom Volume Profiles"
    page.window.width = 410
    page.window.maximizable = False
    page.window.resizable = False
    load_data()  # Load existing profiles

    def update_ui(page=page):
        """
        Updates the GUI of the window.
        :param page: The Flet page object
        """
        scrollable_list.controls.clear()
        load_data()
        page.controls.clear()

        # Add the header and buttons
        top_part = ft.Column([ft.Container(
            padding=ft.padding.all(10),
            content=ft.Row([ft.Text("Custom Audio Profiles", size=18, weight="bold")]),
        )])
        buttons = ft.Row([
            ft.ElevatedButton(text="Add new profile", on_click=open_popup),
            ft.ElevatedButton(text="Github", on_click=open_github),
            ft.ElevatedButton(text="Contact"),
        ])

        # Add components to the page
        page.add(
            ft.Card(
                content=ft.Container(
                    width=500,
                    content=ft.Column([top_part, buttons, ft.Divider()], spacing=0),
                    padding=ft.padding.symmetric(vertical=10)
                )
            )
        )
        page.add(scrollable_list)
        page.update()

    # Event handler for Enter key
    def on_keyboard(e):
        if e.key == "Enter" and not page.dialog.actions[1].disabled:
            close_popup(e)

    page.on_keyboard_event = on_keyboard

    # Define popup for creating new profiles
    def open_popup(e):
        page.dialog = ft.AlertDialog(
            title=ft.Text("Add a new profile"),
            content=ft.Text("Open all the apps, adjust the volume as necessary, then give it a name and press enter"),
            actions=[
                ft.TextField(label="Profile name", on_change=update_button, autofocus=True),
                ft.TextButton("Create", on_click=close_popup, disabled=True)
            ],
        )
        page.dialog.open = True
        page.update()

    # Opens GitHub URL
    def open_github(e):
        webbrowser.open("https://github.com/basilbenny1002/custom_volume_profiles")

    # Updates popup button state
    def update_button(e):
        page.dialog.actions[1].disabled = not e.control.value
        page.update()

    # Closes popup and saves new profile
    def close_popup(e):
        CoInitialize()
        profile_name = page.dialog.actions[0].value
        save_profile(profile_name)
        load_data()
        page.dialog.open = False
        update_ui(page)
        page.update()

    # Add header and scrollable list to page
    top_part = ft.Column([ft.Container(
        padding=ft.padding.all(10),
        content=ft.Row([ft.Text("Custom Audio Profiles", size=18, weight="bold")]),
    )])
    buttons = ft.Row([
        ft.ElevatedButton(text="Add new profile", on_click=open_popup),
        ft.ElevatedButton(text="Github", on_click=open_github),
        ft.ElevatedButton(text="Contact"),
    ])
    page.add(
        ft.Card(
            content=ft.Container(
                width=500,
                content=ft.Column([top_part, buttons, ft.Divider()], spacing=0),
                padding=ft.padding.symmetric(vertical=10)
            )
        )
    )
    page.add(scrollable_list)

    # Update the UI
    page.update()
    update_ui()


if __name__ == "__main__":
    # Run the application
    ft.app(target=main)
