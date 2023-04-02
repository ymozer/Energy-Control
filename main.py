import signal
import win32api
import asyncio
import subprocess
import logging
import time
import win32api as win32
import win32gui
import powerplan


class Monitor:

    def __init__(self, *args, **kwargs):
        self.name = ""
        self.brightness = kwargs.get("brightness")
        self.timeout = kwargs.get("timeout")
        self.monitor = kwargs.get("monitor")
        self.windows = []
        self.powermode = []
        self.input = ""
        self.PROGRAM = "ControlMyMonitor.exe"
        # ANSI colors
        self.c = (
            "\033[0m",   # End of color
            "\033[36m",  # Cyan
            "\033[91m",  # Red
            "\033[35m",  # Magenta
        )

    async def set_brightness(self, monitor, brightness):
        print("Setting brightness of monitor: " + monitor + " to: " + str(brightness))
        completed = subprocess.call(args=[self.PROGRAM, "/SetValue", monitor, "10", str(brightness)], shell=True)
        return completed

    async def get_brightness(self, monitor):
        completed = subprocess.run(args=[self.PROGRAM, "/GetValue", monitor, "10"], shell=True, capture_output=True)
        completed = completed.returncode
        return completed

    # win32api focused window list
    def get_cursor_pos(self):
        return win32.GetCursorPos()

    async def get_power_mode(self):  # not working after first run (not capable of async)
        guid = await powerplan.get_current_scheme_guid()
        if guid == "a1841308-3541-4fab-bc81-f71556f20b4a":
            print("powersaver")
            return 1
        elif guid == "381b4222-f694-41f0-9685-ff5bb260df2e":
            print("balanced")
            return 2
        elif guid == "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c":
            print("high")
            return 3

    async def set_power_mode(self, mode):
        if mode == await self.get_power_mode():
            print(self.c[1]+"Power mode is same. No changes made.")
            return
        elif mode == 1:  # powersaver
            print("Power mode changed to POWERSAVER")
            powerplan.change_current_scheme_to_powersaver()
        elif mode == 2:  # balanced
            print("Power mode changed to BALANCED")
            powerplan.change_current_scheme_to_balanced()
        elif mode == 3:  # high
            print("Power mode changed to PERFORMANCE")
            powerplan.change_current_scheme_to_high()

    def print_window_text(self, hwnd, lparam):
        if win32gui.IsWindowVisible(hwnd):
            self.windows.append(win32gui.GetWindowText(hwnd))

    def get_window_titles(self):
        self.windows.clear()
        win32gui.EnumWindows(self.print_window_text, None)
        return self.windows

    def get_focused_window(self):
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        return (hwnd, title)

    def reset_monitor(self, monitor):
        print("Resetting monitor: " + monitor)
        completed = subprocess.call(args=[self.PROGRAM, "/SetValue", monitor, "04", str(1)], shell=True)
        return completed

    def power(self, monitor):
        print("Switching monitor: " + monitor + " on/off.")
        completed = subprocess.call(args=[self.PROGRAM, "/SwitchValue", monitor, "D6", "1", "4"], shell=True)
        return completed

    def power_on(self, monitor):
        print("Powering on monitor: " + monitor)
        completed = subprocess.call(args=[self.PROGRAM, "/SetValue", monitor, "D6", "1"], shell=True)
        return completed

    def set_input(self, monitor):
        completed = subprocess.call(args=[self.PROGRAM, "/SwitchValue", monitor, "60", "15", "17"], shell=True)
        return completed

    async def getIdleTime(self):
        return (win32api.GetTickCount() - win32api.GetLastInputInfo()) / 1000.0

    async def printIdleTime(self):
        while True:
            print(self.c[1] + "Idle time: " + str(await self.getIdleTime()))
            await asyncio.sleep(1)

    async def blend(self):
        print(self.monitor[0], self.monitor[1], self.brightness, self.timeout)
        while True:
            await asyncio.sleep(1)
            window_titles = list(filter(None,  self.get_window_titles()))
            # print(window_titles)
            # print(self.get_cursor_pos())
            current_power_mode = await self.get_power_mode()
            print(current_power_mode)
            # User is using the computer
            if await self.getIdleTime() == 0 and await self.get_brightness(self.monitor[0]) > 0 and await self.get_brightness(self.monitor[1]) > 0:
                if current_power_mode != 2:
                    await self.set_power_mode(2)
                continue
            # User is not using the computer
            if await self.getIdleTime() > self.timeout and await self.get_brightness(self.monitor[0]) > 0 and await self.get_brightness(self.monitor[1]) > 0:
                await self.set_brightness(self.monitor[0], 0)
                await self.set_brightness(self.monitor[1], 0)
                await self.set_power_mode(1)  # powersaver
            # User came back to the computer
            if await self.getIdleTime() >= 0 and await self.getIdleTime() < 0.5 and await self.get_brightness(self.monitor[0]) == 0 and await self.get_brightness(self.monitor[1]) == 0:
                print(f"Idle time is greater than { self.timeout }")
                await self.set_brightness(self.monitor[0], self.brightness)
                await self.set_brightness(self.monitor[1], self.brightness)
                await self.set_power_mode(2)  # balanced

    async def worker(self):
        # add other functions here to run them concurrently
        futures = [self.blend(), self.printIdleTime()]
        return await asyncio.gather(*futures)

    async def worker_catch(self):
        try:
            return await self.worker()
        except (asyncio.CancelledError, KeyboardInterrupt):
            print('Cancelled task')
        except Exception as ex:
            print('Exception:', ex)
        return None


if __name__ == '__main__':
    # ASUS --> VG249Q1R, N1LMDW000684
    # Lenovo --> LEN P24q-20 , V305C3L8

    monitors = ["V305C3L8", "N1LMDW000684"]

    kwargs = {
        "timeout": 5,
        "brightness": 100,
        "monitor": monitors
    }
    loop = asyncio.get_event_loop()
    loop.set_debug(False)
    # logging.basicConfig(level=logging.DEBUG)
    task = None
    try:
        worker_class = Monitor(**kwargs)
        task = asyncio.ensure_future(worker_class.worker_catch())
        result = loop.run_until_complete(task)

        print('Result: {}'.format(result))
    except KeyboardInterrupt:
        if task:
            print(worker_class.c[0]+'Interrupted, cancelling tasks')
            task.cancel()
            task.exception()
    finally:
        loop.stop()
        loop.close()
