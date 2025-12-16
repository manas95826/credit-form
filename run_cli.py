"""Entry point for running the CLI."""
import sys
from src.cli.main import main

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python run_cli.py <input_file> [output_file]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    main(input_file, output_file)

