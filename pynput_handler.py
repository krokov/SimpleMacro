import time
from pynput import keyboard, mouse

class PynputHandler:
    def __init__(self):
        self.is_recording = False
        self.recorded_events = []
        self.hotkeys = {}
        self.paused = False
        
        # New attributes for the recording delay
        self.recording_delay = 0.1  # Default delay of 100ms
        self.recording_start_time = 0

        self.mouse_listener = mouse.Listener(
            on_click=self.on_click, on_move=self.on_move, on_scroll=self.on_scroll)
        self.keyboard_listener = keyboard.Listener(
            on_press=self.on_press, on_release=self.on_release)
        
        self.mouse_listener.start()
        self.keyboard_listener.start()

    def set_hotkeys(self, hotkeys_map):
        self.hotkeys = hotkeys_map
        
    def set_recording_delay(self, delay):
        """Allows the main app to set the recording delay."""
        try:
            self.recording_delay = float(delay)
        except (ValueError, TypeError):
            self.recording_delay = 0.1 # Fallback to default if value is invalid

    def _add_event(self, event_type, *args):
        # Only add the event if recording is active AND the delay period has passed.
        if self.is_recording and time.time() > self.recording_start_time:
            event = {'time': time.time(), 'type': event_type, 'data': args}
            self.recorded_events.append(event)

    def on_press(self, key):
        if self.paused: return
        if key in self.hotkeys:
            self.hotkeys[key]()
            return
        self._add_event('key_press', str(key))

    def on_release(self, key):
        self._add_event('key_release', str(key))

    def on_click(self, x, y, button, pressed):
        self._add_event('mouse_click', x, y, str(button), pressed)

    def on_move(self, x, y):
        # Check the time here as well to avoid spamming move events during the delay
        if self.is_recording and time.time() > self.recording_start_time and self.recorded_events:
            last_event = self.recorded_events[-1]
            if last_event['type'] == 'mouse_move':
                if abs(x - last_event['data'][0]) < 5 and abs(y - last_event['data'][1]) < 5:
                    return
        self._add_event('mouse_move', x, y)
    
    def on_scroll(self, x, y, dx, dy):
        self._add_event('mouse_scroll', x, y, dx, dy)
    
    def play_macro(self, events):
        if not events: return
        keyboard_controller, mouse_controller = keyboard.Controller(), mouse.Controller()
        start_time = events[0]['time']
        for event in events:
            delay = event['time'] - start_time
            time.sleep(max(0, delay))
            start_time = event['time']
            event_type, data = event['type'], event['data']
            try:
                if event_type == 'key_press':
                    key = eval(f"keyboard.{data[0]}") if data[0].startswith('Key.') else data[0].strip("'")
                    keyboard_controller.press(key)
                elif event_type == 'key_release':
                    key = eval(f"keyboard.{data[0]}") if data[0].startswith('Key.') else data[0].strip("'")
                    keyboard_controller.release(key)
                elif event_type == 'mouse_move':
                    mouse_controller.position = (int(data[0]), int(data[1]))
                elif event_type == 'mouse_click':
                    x, y, button_str, pressed = data
                    mouse_controller.position = (int(x), int(y))
                    button = eval(f"mouse.{button_str}")
                    if pressed: mouse_controller.press(button)
                    else: mouse_controller.release(button)
                elif event_type == 'mouse_scroll':
                    x, y, dx, dy = data
                    mouse_controller.position = (int(x), int(y))
                    mouse_controller.scroll(int(dx), int(dy))
            except Exception as e:
                print(f"Error playing back event {event}: {e}")

    def start_recording(self):
        self.recorded_events = []
        self.is_recording = True
        # Set the time when recording is allowed to actually start
        self.recording_start_time = time.time() + self.recording_delay

    def stop_recording(self):
        self.is_recording = False
        return self.recorded_events

    def stop_listeners(self):
        self.keyboard_listener.stop()
        self.mouse_listener.stop()

    def pause(self): self.paused = True
    def resume(self): self.paused = False