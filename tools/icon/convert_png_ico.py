from pathlib import Path

from PIL import Image

here = Path(__file__).resolve().parent
repo_root = here.parents[1]

candidates = (
    here / "input.png",
    repo_root / "tools" / "icon" / "input.png",
)

src = next((p for p in candidates if p.exists()), None)
if src is None:
    tried = ", ".join(str(p) for p in candidates)
    raise FileNotFoundError(f"Could not find input.png. Tried: {tried}")

out = here / "output.ico"

img = Image.open(src).convert("RGBA")
sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
img.save(out, format="ICO", sizes=sizes)

print(f"Created: {out}")
print(f"Source: {src}")
