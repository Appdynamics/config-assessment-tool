import logging
import sys

# notice: NO 'run_plugin' function is defined here.

def main():
    logging.basicConfig(level=logging.INFO)
    logging.info("Demo Standalone CLI: I am running!")
    print("Demo Standalone CLI: This output is visible in CLI.")

    if len(sys.argv) > 1:
        print(f"Demo Standalone CLI: Received args: {sys.argv[1:]}")

if __name__ == "__main__":
    main()
