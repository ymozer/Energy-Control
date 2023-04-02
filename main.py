import win32api
import asyncio
import subprocess
import logging
import time
import win32api as win32
import win32con
import shlex
import os


class Monitor:

    def __init__(self, *args, **kwargs):
        self.name = ""
        self.brightness = kwargs.get("brightness")
        self.timeout = kwargs.get("timeout")
        self.monitor = kwargs.get("monitor")
        self.power = False
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
    
    async def focus(self, monitor):
        print("Focusing monitor: " + monitor)
        completed = subprocess.call(args=[self.PROGRAM, "/SwitchValue", monitor, "60", "15", "17"], shell=True)
        return completed

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
            # User is using the computer
            if await self.getIdleTime() == 0 and await self.get_brightness(self.monitor[0]) > 0 and await self.get_brightness(self.monitor[1]) > 0:
                continue
            # User is not using the computer
            if await self.getIdleTime() > self.timeout and await self.get_brightness(self.monitor[0]) > 0 and await self.get_brightness(self.monitor[1]) > 0:
                await self.set_brightness(self.monitor[0], 0)
                await self.set_brightness(self.monitor[1], 0)
            # User came back to the computer
            if await self.getIdleTime() >= 0 and await self.getIdleTime() < 0.5 and await self.get_brightness(self.monitor[0]) == 0 and await self.get_brightness(self.monitor[1]) == 0:
                print(f"Idle time is greater than { self.timeout }")
                await self.set_brightness(self.monitor[0], self.brightness)
                await self.set_brightness(self.monitor[1], self.brightness)


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
    loop.set_debug(True)
    # logging.basicConfig(level=logging.DEBUG)
    task = None
    try:
        worker_class = Monitor(**kwargs)
        task = asyncio.ensure_future(worker_class.worker_catch())
        result = loop.run_until_complete(task)
        print('Result: {}'.format(result))
    except KeyboardInterrupt:
        if task:
            print('Interrupted, cancelling tasks')
            task.cancel()
            task.exception()
    finally:
        loop.close()
