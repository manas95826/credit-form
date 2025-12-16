"""Legacy CLI entry point - uses new architecture."""
from src.cli.main import main
import sys

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python app.py <input_file> [output_file]")
        print("Note: This is a legacy entry point. Consider using: python run_cli.py")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    main(input_file, output_file)
