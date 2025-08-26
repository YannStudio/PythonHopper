import os
import pandas as pd
from helpers import _build_file_index

def test_all_required_exts(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    # PN1 has both pdf and stp
    (src / "PN1.pdf").write_text("pdf")
    (src / "PN1.stp").write_text("step")
    # PN2 only pdf
    (src / "PN2.pdf").write_text("pdf")
    df = pd.DataFrame({"PartNumber": ["PN1", "PN2"]})
    idx = _build_file_index(str(src), [".pdf", ".step", ".stp"])
    groups = [{".step", ".stp"}, {".pdf"}]
    statuses = []
    for pn in df["PartNumber"]:
        hits = idx.get(pn, [])
        hit_exts = {os.path.splitext(h)[1].lower() for h in hits}
        all_present = all(any(ext in hit_exts for ext in g) for g in groups)
        statuses.append(all_present)
    assert statuses == [True, False]
