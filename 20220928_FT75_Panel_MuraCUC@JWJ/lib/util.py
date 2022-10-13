from abc import ABCMeta, abstractmethod
import logging, time, functools, io, pathlib
from typing import List, TypeVar, Any, Callable
Path = TypeVar('Path', str, bytes)
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')

def timer(func:Callable) -> Callable:
    @functools.wraps(func)
    def inner(*args:Any, **kwargs:Any) -> Any:
        b:float = time.time()
        res:Any = func(*args, **kwargs) 
        e:float = time.time()
        logging.info(f'time lapsed(s) : {e-b:.2f}')
        return res
    return inner

class Worker(metaclass=ABCMeta):
    @abstractmethod
    def sort_logs(self, src_folder:Path, file_ext:str, dst_file:Path) -> None:
        raise NotImplementedError()

    @abstractmethod
    def extract(self, src_file:Path, dst:io.TextIOWrapper) -> int:
        raise NotImplementedError()

class LogWorker(Worker):
    @timer
    def sort_logs(self, src_folder: Path, file_ext:str, dst_file: Path) -> None:
        """sort valuable data from a given root folder

        Args:
            src_folder (Path): a given folder
            file_ext (str): file extension
            dst_file (Path): file path stores result
        """
        n:int = 0
        file:pathlib.Path = None
        logging.info('started..')
        with open(dst_file, 'a') as dst:
            for file in pathlib.Path(src_folder).glob(f'*.{file_ext}'):
                n += self.extract(file.absolute(), dst)
        logging.info(f'suspected NG = {n}')