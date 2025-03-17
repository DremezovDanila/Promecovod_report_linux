from pathlib import Path
import os

#  Put root directory in decentralized method.
ROOT_DIR = Path(__file__).resolve().parents[0]
#  LiteralString format.
root_dir = "/".join(os.path.abspath(__file__).split("/")[:-1])
