from Interface import Interface
import Guard
import sys


class NullWriter:
    def write(self, data):
        pass


# Override stdout and stderr with NullWriter in GUI --noconsole mode
# This allow to avoid a bug where tqdm try to write on NoneType
if sys.stdout is None:
    sys.stdout = NullWriter()

if sys.stderr is None:
    sys.stderr = NullWriter()

if __name__ == "__main__":
    if Guard.check():
        program = Interface()
        program.run()
