#!/usr/bin/env python3
"""
Complete PDF Overlay Processing Pipeline
Runs the entire process: build overlay → add colors → optionally add configurations

Usage:
  python run_complete_overlay.py old.pdf new.pdf output.pdf [--colors-only] [--fit]
"""

import argparse
import logging
import subprocess
import sys
import shutil
from pathlib import Path
from typing import Tuple

LOGGER = logging.getLogger("complete_overlay")

SCRIPT_DIR = Path(__file__).resolve().parent

def run_command(cmd: list, description: str) -> bool:
    """Run a command and return success status."""
    LOGGER.info(f"🔄 {description}")
    LOGGER.info(f"   Command: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        if result.stdout:
            LOGGER.info(f"   Output: {result.stdout.strip()}")
        LOGGER.info(f"✅ {description} - SUCCESS")
        return True
    except subprocess.CalledProcessError as e:
        LOGGER.error(f"❌ {description} - FAILED")
        LOGGER.error(f"   Exit code: {e.returncode}")
        if e.stdout:
            LOGGER.error(f"   Stdout: {e.stdout}")
        if e.stderr:
            LOGGER.error(f"   Stderr: {e.stderr}")
        return False

def validate_input_files(old_pdf: Path, new_pdf: Path) -> bool:
    """Validate that input PDF files exist."""
    if not old_pdf.exists():
        LOGGER.error(f"❌ Old PDF file not found: {old_pdf}")
        return False
    if not new_pdf.exists():
        LOGGER.error(f"❌ New PDF file not found: {new_pdf}")
        return False
    LOGGER.info(f"✅ Input files validated:")
    LOGGER.info(f"   Old PDF: {old_pdf} ({old_pdf.stat().st_size:,} bytes)")
    LOGGER.info(f"   New PDF: {new_pdf} ({new_pdf.stat().st_size:,} bytes)")
    return True

def parse_color(color_str: str) -> Tuple[int, int, int]:
    """Parse color string in format 'R,G,B'."""
    try:
        r, g, b = map(int, color_str.split(','))
        if not all(0 <= c <= 255 for c in [r, g, b]):
            raise ValueError("Color values must be 0-255")
        return r, g, b
    except Exception as e:
        raise ValueError(f"Invalid color format '{color_str}'. Use R,G,B format (e.g., '255,0,0')") from e

def main():
    parser = argparse.ArgumentParser(
        description="Complete PDF overlay processing pipeline",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
        epilog="""
Examples:
  # Basic usage with colors only (most reliable):
  python run_complete_overlay.py "old.pdf" "new.pdf" "output.pdf" --colors-only
  
  # Custom colors:
  python run_complete_overlay.py old.pdf new.pdf output.pdf --old-color 255,100,0 --new-color 0,200,100 --colors-only
  
  # Keep intermediate files for debugging:
  python run_complete_overlay.py old.pdf new.pdf output.pdf --colors-only --keep-intermediate --verbose
        """
    )
    
    # Required arguments
    parser.add_argument("old_pdf", type=Path, help="Path to old/original PDF file")
    parser.add_argument("new_pdf", type=Path, help="Path to new/revised PDF file")
    parser.add_argument("output_pdf", type=Path, help="Path to final output PDF")
    
    # Color and appearance options
    parser.add_argument("--old-color", type=str, default="255,0,0", 
                       help="Color for old layer in R,G,B format")
    parser.add_argument("--new-color", type=str, default="0,255,0", 
                       help="Color for new layer in R,G,B format")
    parser.add_argument("--alpha", type=float, default=0.6, 
                       help="Alpha transparency for colored layers (0.0-1.0)")
    
    # Processing options
    parser.add_argument("--fit", action="store_true", 
                       help="Fit pages to match size during overlay build")
    parser.add_argument("--keep-intermediate", action="store_true", 
                       help="Keep intermediate files for debugging")
    parser.add_argument("--colors-only", action="store_true",
                       help="Only add colors, skip layer configurations (recommended)")
    
    # Output options
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging")
    parser.add_argument("--quiet", action="store_true", help="Suppress all output except errors")
    
    args = parser.parse_args()
    
    # Set up logging
    if args.quiet:
        level = logging.ERROR
    elif args.verbose:
        level = logging.INFO
    else:
        level = logging.WARNING
    
    logging.basicConfig(
        level=level, 
        format="%(levelname)s: %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)]
    )
    
    try:
        # Validate inputs
        LOGGER.info("🚀 Starting Complete PDF Overlay Processing Pipeline")
        LOGGER.info("=" * 60)
        
        if not validate_input_files(args.old_pdf, args.new_pdf):
            sys.exit(1)
        
        # Parse colors
        try:
            old_color = parse_color(args.old_color)
            new_color = parse_color(args.new_color)
            LOGGER.info(f"   Old layer color: RGB{old_color}")
            LOGGER.info(f"   New layer color: RGB{new_color}")
            LOGGER.info(f"   Alpha transparency: {args.alpha}")
        except ValueError as e:
            LOGGER.error(f"❌ {e}")
            sys.exit(1)
        
        if not (0.0 <= args.alpha <= 1.0):
            LOGGER.error("❌ Alpha must be between 0.0 and 1.0")
            sys.exit(1)
        
        # Define intermediate file paths
        base_name = args.output_pdf.stem
        output_dir = args.output_pdf.parent
        overlay_file = output_dir / f"{base_name}_overlay.pdf"
        colored_file = output_dir / f"{base_name}_colored.pdf"
        
        LOGGER.info(f"📁 Working directory: {output_dir}")
        LOGGER.info(f"📄 Final output: {args.output_pdf}")
        
        # Ensure output directory exists
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Step 1: Build overlay with three layers
        LOGGER.info("\n" + "=" * 60)
        LOGGER.info("STEP 1: Building Three-Layer Overlay")
        LOGGER.info("=" * 60)
        
        build_script = SCRIPT_DIR / "overlay_build.py"
        build_cmd = [
            sys.executable,
            str(build_script),
            str(args.old_pdf),
            str(args.new_pdf), 
            str(overlay_file)
        ]
        if args.fit:
            build_cmd.append("--fit")
        if args.verbose:
            build_cmd.append("--verbose")
            
        if not run_command(build_cmd, "Building three-layer overlay"):
            LOGGER.error("❌ Failed to build overlay")
            sys.exit(1)
        
        if not overlay_file.exists():
            LOGGER.error(f"❌ Overlay file was not created: {overlay_file}")
            sys.exit(1)
        
        # Step 2: Add colors to layers
        LOGGER.info("\n" + "=" * 60)
        LOGGER.info("STEP 2: Adding Colors to Layers")
        LOGGER.info("=" * 60)
        
        layerize_script = SCRIPT_DIR / "overlay_layerize.py"
        color_cmd = [
            sys.executable,
            str(layerize_script),
            str(overlay_file),
            str(colored_file),
            "--old-color", args.old_color,
            "--new-color", args.new_color,
            "--old-alpha", str(args.alpha),
            "--new-alpha", str(args.alpha)
        ]
        if args.verbose:
            color_cmd.append("--verbose")
            
        if not run_command(color_cmd, "Adding colors to layers"):
            LOGGER.error("❌ Failed to add colors")
            sys.exit(1)
        
        if not colored_file.exists():
            LOGGER.error(f"❌ Colored file was not created: {colored_file}")
            sys.exit(1)
        
        # Step 3: Copy or add configurations
        if args.colors_only:
            LOGGER.info("\n🔧 Colors-only mode: Copying colored PDF to final output")
            shutil.copy2(colored_file, args.output_pdf)
            LOGGER.info(f"✅ Final PDF ready: {args.output_pdf}")
        else:
            LOGGER.info("\n⚠️  Advanced configurations not implemented in this version")
            LOGGER.info("   Using --colors-only mode instead for reliability")
            shutil.copy2(colored_file, args.output_pdf)
            LOGGER.info(f"✅ Final PDF ready: {args.output_pdf}")
        
        if not args.output_pdf.exists():
            LOGGER.error(f"❌ Final output file was not created: {args.output_pdf}")
            sys.exit(1)
        
        # Clean up intermediate files if requested
        if not args.keep_intermediate:
            LOGGER.info("\n🧹 Cleaning up intermediate files...")
            for temp_file in [overlay_file, colored_file]:
                if temp_file.exists():
                    temp_file.unlink()
                    LOGGER.info(f"   Removed: {temp_file}")
        else:
            LOGGER.info(f"\n📁 Intermediate files kept:")
            LOGGER.info(f"   Overlay: {overlay_file}")
            LOGGER.info(f"   Colored: {colored_file}")
        
        # Final success summary
        LOGGER.info("\n" + "🎉" * 20)
        LOGGER.info("SUCCESS: Complete PDF Overlay Processing Pipeline")
        LOGGER.info("🎉" * 20)
        
        final_size = args.output_pdf.stat().st_size
        LOGGER.info(f"📄 Final output: {args.output_pdf}")
        LOGGER.info(f"📊 File size: {final_size:,} bytes")
        LOGGER.info(f"🎨 Colors: Old=RGB{old_color}, New=RGB{new_color}, Current=Black")
        LOGGER.info(f"🔧 Alpha: {args.alpha}")
        LOGGER.info("📋 Three Layers Available:")
        LOGGER.info("   • Old p1 (Red)")
        LOGGER.info("   • New p1 (Green)")  
        LOGGER.info("   • Current p1 (Black)")
        LOGGER.info("\n✅ Ready to view in PDF reader with layer support!")
        
    except KeyboardInterrupt:
        LOGGER.error("\n❌ Process interrupted by user")
        sys.exit(1)
    except Exception as e:
        LOGGER.error(f"\n❌ Unexpected error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
