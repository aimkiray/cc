from pystray import MenuItem as item, Icon as icon, Menu as menu
import threading
from PIL import Image

def create_tray_icon(image_path, title, show_window, hide_window, exit_app):
    """Create a system tray icon."""
    window_visible = True

    def toggle_window_visibility(icon, item):
        nonlocal window_visible
        if window_visible:
            hide_window()
        else:
            show_window()
        window_visible = not window_visible  # Toggle visibility state
        icon.update_menu()

    image = Image.open(image_path)
    icon_menu = menu(
        item('显示', lambda icon, item: toggle_window_visibility(icon, item), visible=lambda item: not window_visible, default=True),
        item('隐藏', lambda icon, item: toggle_window_visibility(icon, item), visible=lambda item: window_visible),
        item('退出', lambda icon, item: exit_app())
    )

    tray_icon = icon(title, image, title, icon_menu)
    tray_icon.run()

def start_tray_icon(image_path, title, show_window, hide_window, exit_app):
    """Start the tray icon in a separate thread."""
    tray_thread = threading.Thread(target=create_tray_icon, args=(image_path, title, show_window, hide_window, exit_app))
    tray_thread.daemon = True
    tray_thread.start()
    return tray_thread