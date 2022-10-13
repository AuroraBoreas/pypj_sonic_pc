"""
# this is a simple module that identifies suspected NG samples from tact log files.

## author
- 20220928 @ZL

## changelog
- v0.01: initial build
- v0.02: algorithm optimization
- v0.03: template pattern

"""

import io, sys, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from lib.util import logging, Path, List, LogWorker

class TLogWorker(LogWorker):
    def extract(self, src_file:Path, dst:io.TextIOWrapper) -> int:
        idn_mura_cuc_begin:str = '0_Diag_TAG_OOB_Tact'
        nrow_mura_cuc_adj:str  = 'MURA_CUC'
        ncol_scc:int = 6
        ncol_lin:int = 7
        ncol_res:int = 10
        res_fmt:str  = '{},{},{},{},{}'
        fn:str       = os.path.split(src_file)[-1]
        n:int        = 0

        with open(src_file, 'r', encoding='utf-8') as src:
            lines:List  = src.readlines()
            max_idx:int = len(lines) - 1
            for i, line in enumerate(lines):
                res:List = line.strip().split(',')
                if res[ncol_res] == idn_mura_cuc_begin:
                    if (i + 1) <= max_idx:
                        if nrow_mura_cuc_adj in lines[i+1]:
                            # logging.info(res_fmt.format(fn, res[ncol_scc], res[ncol_lin], lines[i+1].split(',')[ncol_res], 'OK'))
                            dst.writelines(res_fmt.format(fn, res[ncol_scc], res[ncol_lin], lines[i+1].split(',')[ncol_res], 'OK\n'))
                        else:
                            # logging.info(res_fmt.format(fn, res[ncol_scc], res[ncol_lin], lines[i+1].split(',')[ncol_res], 'NG'))
                            dst.writelines(res_fmt.format(fn, res[ncol_scc], res[ncol_lin], lines[i+1].split(',')[ncol_res], 'NG\n'))
                            n += 1
        return n

def main() -> None:
    w:LogWorker     = TLogWorker()
    src_folder:Path = 'data'
    w.sort_logs(src_folder, 'csv', 'result.csv')

if __name__ == '__main__':
    main()