from screeninfo import get_monitors

monitors = get_monitors()

for monitor in monitors:
    print(f"Monitor: {monitor.name}, Resolução: {monitor.width}x{monitor.height}")