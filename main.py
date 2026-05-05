import os
import time
import win32api
import win32gui
import win32com.client
import win32con
from tkinter import Tk, Toplevel, Label, filedialog
from PIL import Image, ImageTk
import traceback
import ctypes

# ---------------------------
# INITIALIZATION (DPI Awareness)
# ---------------------------
try:
    # Try to set per-monitor DPI awareness (Windows 8.1+)
    # 2 = PROCESS_PER_MONITOR_DPI_AWARE
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        # Fallback for older Windows 10 or Windows 7/8
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass



# ---------------------------
# MONITOR UTILITIES
# ---------------------------
def get_monitors():
    """Returns a list of monitor rectangles (left, top, right, bottom)."""
    try:
        monitors = [m[2] for m in win32api.EnumDisplayMonitors()]
        monitors.sort(key=lambda x: x[0])
        return monitors
    except Exception as e:
        print(f"Error enumerating monitors: {e}")
        return []


def get_slideshow_monitor(monitors):
    """Detects which monitor the PowerPoint slideshow is currently on."""
    # PowerPoint slideshow window class is "ScreenClass"
    # In Windows 10, there might be multiple windows with this class if Presenter View is used.
    def callback(hwnd, extra):
        if win32gui.GetClassName(hwnd) == "ScreenClass" and win32gui.IsWindowVisible(hwnd):
            extra.append(hwnd)
        return True

    hwnds = []
    try:
        win32gui.EnumWindows(callback, hwnds)
    except Exception as e:
        print(f"Error enumerating windows: {e}")
        # Fallback to simple FindWindow
        h = win32gui.FindWindow("ScreenClass", None)
        if h: hwnds = [h]

    if not hwnds:
        return None

    # Find the window with the largest area (the actual slideshow, not presenter thumbnails)
    best_monitor = None
    max_area = -1
    
    for hwnd in hwnds:
        try:
            # Use GetWindowRect. With DPI awareness set, this should return 
            # coordinates consistent with EnumDisplayMonitors.
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            area = width * height
            
            # Use center of the window
            cx = (rect[0] + rect[2]) // 2
            cy = (rect[1] + rect[3]) // 2

            for m in monitors:
                if m[0] <= cx < m[2] and m[1] <= cy < m[3]:
                    if area > max_area:
                        max_area = area
                        best_monitor = m
        except Exception as e:
            print(f"Error checking window {hwnd}: {e}")
            continue
    
    return best_monitor


def choose_preview_monitor(monitors):
    """Selects a monitor for the preview window, avoiding the slideshow monitor."""
    if not monitors:
        return None
    if len(monitors) == 1:
        return None

    slideshow = get_slideshow_monitor(monitors)
    if slideshow:
        others = [m for m in monitors if m != slideshow]
        # Prefer the monitor to the right, or just the largest/next one
        return max(others, key=lambda m: m[0]) if others else None

    # Fallback: use the second monitor if available
    return monitors[-1]


# ---------------------------
# PREVIEW WINDOW
# ---------------------------
class PreviewWindow:
    def __init__(self, root, monitor):
        self.win = Toplevel(root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.config(bg="black")

        self.x1, self.y1, self.x2, self.y2 = monitor
        self.width = self.x2 - self.x1
        self.height = self.y2 - self.y1
        
        self.win.geometry(f"{self.width}x{self.height}+{self.x1}+{self.y1}")

        self.label = Label(self.win, bg="black")
        self.label.pack(expand=True, fill="both")
        
        # Initial black image
        self.clear()

    def clear(self):
        self.label.config(image="", text="Waiting for slide...", fg="white")
        self.label.image = None

    def show_image(self, path):
        if not os.path.exists(path):
            return

        try:
            # Use a context manager to ensure the file is closed quickly
            with Image.open(path) as img:
                # Calculate aspect ratio to fit without stretching
                img_ratio = img.width / img.height
                win_ratio = self.width / self.height

                if img_ratio > win_ratio:
                    new_w = self.width
                    new_h = int(new_w / img_ratio)
                else:
                    new_h = self.height
                    new_w = int(new_h * img_ratio)

                img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                self.label.config(image=photo, text="")
                self.label.image = photo
        except Exception as e:
            print(f"Error showing image: {e}")


# ---------------------------
# POWERPOINT CONTROL
# ---------------------------
class PowerPointController:
    def __init__(self, file_path):
        self.app = None
        self.pres = None
        self.ssw = None
        
        abs_path = os.path.abspath(file_path).replace("/", "\\")
        
        try:
            self.app = win32com.client.Dispatch("PowerPoint.Application")
            self.app.Visible = True  # Ensure PowerPoint is visible to avoid frame error
            self.pres = self.app.Presentations.Open(abs_path, WithWindow=True)
            self.ssw = self.pres.SlideShowSettings.Run()
        except Exception as e:
            print(f"Failed to start PowerPoint: {e}")
            raise

    def get_current_slide_index(self):
        try:
            if self.is_running():
                return self.ssw.View.CurrentShowPosition
        except:
            pass
        return -1

    def export_slide(self, index, path):
        try:
            if 1 <= index <= self.pres.Slides.Count:
                # Remove old file first to avoid conflicts
                if os.path.exists(path):
                    try:
                        os.remove(path)
                    except:
                        pass
                
                # PowerPoint can sometimes be busy; try a few times
                for _ in range(3):
                    try:
                        self.pres.Slides(index).Export(path, "PNG")
                        if os.path.exists(path):
                            return True
                    except:
                        time.sleep(0.1)
                
                return False
        except Exception as e:
            print(f"Error exporting slide {index}: {e}")
        return False

    def is_running(self):
        try:
            # Also check if the window handle still exists
            return self.app.SlideShowWindows.Count > 0
        except:
            return False

    def close(self):
        try:
            if self.pres:
                self.pres.Saved = True  # Prevent "Save changes?" prompt
                self.pres.Close()
            if self.app:
                self.app.Quit()
        except:
            pass


# ---------------------------
# MAIN APP
# ---------------------------
class PPTPreviewApp:
    def __init__(self, root, file_path):
        self.root = root
        self.temp_img = os.path.join(os.environ.get("TEMP", "."), "ppt_prev_cache.png")
        self.last_slide = -1
        self.preview = None
        
        print(f"Opening: {file_path}")
        self.ppt = PowerPointController(file_path)
        
        # Wait for the slideshow window to be detected (up to 5 seconds)
        print("Waiting for slideshow to initialize...")
        preview_monitor = None
        for _ in range(10):  # 10 * 0.5s = 5s
            time.sleep(0.5)
            monitors = get_monitors()
            if not monitors: continue
            
            preview_monitor = choose_preview_monitor(monitors)
            # If we found the slideshow and picked a DIFFERENT monitor for preview, we are good.
            # If choose_preview_monitor returned a monitor that is NOT the one the slideshow is on,
            # it means detection worked.
            slideshow_mon = get_slideshow_monitor(monitors)
            if slideshow_mon and preview_monitor and preview_monitor != slideshow_mon:
                break
            
            # If no slideshow found yet, keep waiting
            if not slideshow_mon:
                continue
            
            # If slideshow found but no other monitor, we can't do much
            if slideshow_mon and len(monitors) == 1:
                break

        if preview_monitor:
            print(f"Preview monitor detected: {preview_monitor}")
            self.preview = PreviewWindow(self.root, preview_monitor)
        else:
            print("Could not find a suitable secondary monitor for preview.")

        self.running = True
        # Local bindings for when the Python windows have focus
        self.root.bind("<Escape>", self.stop)
        if self.preview:
            self.preview.win.bind("<Escape>", self.stop)

    def stop(self, event=None):
        print("Stop requested.")
        self.running = False

    def run(self):
        try:
            print("Monitoring started. Press ESC to stop.")
            while self.running:
                self.root.update()
                
                # GLOBAL ESCAPE DETECTION
                # 0x1B is the virtual key code for ESCAPE
                if win32api.GetAsyncKeyState(0x1B) & 0x8000:
                    print("Global Escape detected. Stopping...")
                    break

                if not self.ppt.is_running():
                    print("Slideshow ended (detected via COM).")
                    break

                current = self.ppt.get_current_slide_index()

                if current != self.last_slide:
                    print(f"Slide changed: {self.last_slide} -> {current}")
                    
                    if current == 1:
                        if self.preview:
                            self.preview.clear()
                    
                    elif current > 1 and self.preview:
                        if self.ppt.export_slide(current - 1, self.temp_img):
                            self.preview.show_image(self.temp_img)

                    self.last_slide = current

                time.sleep(0.1)  # Increased responsiveness
        except Exception as e:
            print(f"Runtime error: {e}")
            traceback.print_exc()
        finally:
            self.cleanup()

    def cleanup(self):
        print("Cleaning up...")
        self.running = False
        if self.ppt:
            self.ppt.close()
        if os.path.exists(self.temp_img):
            try:
                os.remove(self.temp_img)
            except:
                pass
        try:
            self.root.destroy()
        except:
            pass


# ---------------------------
# ENTRY POINT
# ---------------------------
def main():
    # Attempt to extend display only if needed
    monitors = win32api.EnumDisplayMonitors()
    if len(monitors) < 2:
        print("Setting display to Extend mode...")
        os.system("DisplaySwitch.exe /extend")
        time.sleep(3)
    else:
        print("Display already extended or multiple monitors detected.")

    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select PowerPoint Presentation",
        filetypes=[("PowerPoint Files", "*.pptx *.ppt")]
    )

    if file_path:
        try:
            app = PPTPreviewApp(root, file_path)
            app.run()
        except Exception as e:
            print(f"Application failed to start: {e}")
            traceback.print_exc()
    else:
        print("No file selected.")

    print("Application exited.")


if __name__ == "__main__":
    main()