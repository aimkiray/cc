from pystray import MenuItem as item, Icon as icon, Menu as menu
import threading
from PIL import Image

def create_tray_icon(window):
    """Create a system tray icon."""

    image = Image.open(window.icon)
    icon_menu = menu(
        item('显示', lambda _: window.show_window(), visible=lambda _: not window.visible, default=lambda _: not window.visible),
        item('隐藏', lambda _: window.hide_window(), visible=lambda _: window.visible, default=lambda _: window.visible),
        item('退出', lambda _: window.exit_app())
    )

    tray_icon = icon(window.title, image, window.title, icon_menu)
    window.tray = tray_icon
    tray_icon.run()

def start_tray_icon(window):
    """Start the tray icon in a separate thread."""
    tray_thread = threading.Thread(target=create_tray_icon, args=(window,))
    tray_thread.daemon = True
    tray_thread.start()
    return tray_thread