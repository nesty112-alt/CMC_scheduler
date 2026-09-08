from datetime import datetime
from pathlib import Path

OUTPUT = Path(__file__).resolve().parent / "output" / "python_result.txt"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)
stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
OUTPUT.write_text(f"CMC_SAMPLE_OK {stamp}\n", encoding="utf-8")
print(f"CMC_SAMPLE_OK {stamp}")
