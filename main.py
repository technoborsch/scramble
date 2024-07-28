from Interface import Interface
import Guard

if __name__ == "__main__":
    if Guard.check():
        program = Interface()
        program.run()
