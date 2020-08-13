import os, sys
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from lib.pkg import runbat

base_dir = os.path.dirname(__file__)
bat1 = os.path.join(base_dir, r"bats\STEP1 create_files.cmd")
bat2 = os.path.join(base_dir, r"bats\STEP2 rename_files.cmd")
bat3 = os.path.join(base_dir, r"bats\STEP3 delete_files.cmd")
bats = [bat1, bat2, bat3]
for bat in bats:
    runbat.run_batch(bat)