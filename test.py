import sys
import time

for n in range(100):
    time.sleep(1)
    if n == 99:
        sys.stdout.write('completed')
    else:
        sys.stdout.write(str(n) + ' ')
    sys.stdout.flush()