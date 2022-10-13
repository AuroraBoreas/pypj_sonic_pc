"""
# this is a simple module that identifies suspected NG samples from jig log files.

## author
- 20220928 @ZL

## changelog
- v0.01: initial build
- v0.02: bugfix
- v0.03: algorithm optimization
- v0.04: template pattern

"""

import os, io
from lib.util import Path, logging, List, LogWorker

class JLogWorker(LogWorker):
    def extract(self, src_file:Path, dst:io.TextIOWrapper) -> int:
        # panel scc x1
        idn_panel_scc:str = "diag_ffa_read result OK: READ DATA="
        idn_panel_0:str   = "diag_ffa_read result OK: READ DATA=0"
        # panel ID
        idn_panel_id:str  = "diag_panel_id_read result OK: PanelID="
        nrow_pid:int      = 16
        # mura verify result x1
        nrow_mura_fmt_01:int = 32
        nrow_mura_fmt_02:int = 164
        idn_mura_verify:str  = "diag_verify_mura_data"
        # cuc  verify result x1
        idn_cuc_verify:str   = "diag_verify_cuc_data"
        nrow_cuc_fmt_01:int  = 50
        nrow_cuc_fmt_02:int  = 304
        res:str = None
        delimiter:str = '] '
        n:int = 0

        with open(src_file, 'r') as src:
            src_lines:List = src.readlines()
            for i, line in enumerate(src_lines):
                res = line.strip().split(delimiter)[-1]
                if res != idn_panel_0 and res.startswith(idn_panel_scc):
                    if i + nrow_pid < len(src_lines) and i + nrow_cuc_fmt_01 < len(src_lines) and i + nrow_mura_fmt_01 < len(src_lines):
                        log_res:str       = '{},{},{}'.format(os.path.split(src_file)[-1], res, src_lines[i + nrow_pid].strip().split(idn_panel_id)[-1])
                        mura_res:bool     = idn_mura_verify in src_lines[i + nrow_mura_fmt_01]
                        cuc_res:bool      = idn_cuc_verify in src_lines[i + nrow_cuc_fmt_01]
                        mura_cuc_res:List = [mura_res, cuc_res]
                        if not mura_res:
                            if i + nrow_cuc_fmt_02 < len(src_lines) and i + nrow_mura_fmt_02 < len(src_lines):
                                mura_res     = idn_mura_verify in src_lines[i + nrow_mura_fmt_02]
                                cuc_res      = idn_cuc_verify in src_lines[i + nrow_cuc_fmt_02]
                                mura_cuc_res = [mura_res, cuc_res]
                    else:
                        log_res = 'None'
                        mura_cuc_res = [False, False]
                    if log_res == 'None':
                        # logging.info(log_res + ' -> None')
                        dst.write('{},{}\n'.format(log_res, 'None'))
                    else:
                        if all(mura_cuc_res):
                            # logging.info(log_res + ' -> OK')
                            dst.write('{},{}\n'.format(log_res, 'OK'))
                        else:
                            # logging.info(log_res + ' -> NG')
                            dst.write('{},{}\n'.format(log_res, 'NG'))
                            n += 1
        return n

def main() -> None:
    w:LogWorker     = JLogWorker()
    src_folder:Path = 'data'
    w.sort_logs(src_folder, 'log', 'result.csv')

if __name__ == '__main__':
    main()