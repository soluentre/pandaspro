from abc import ABCMeta, abstractmethod


class cpdReport(metaclass=ABCMeta):
    @classmethod
    @abstractmethod
    def data(cls):
        pass

    @classmethod
    @abstractmethod
    def cpdbase_engine(cls):
        pass


