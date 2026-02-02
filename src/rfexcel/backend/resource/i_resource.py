from abc import ABC, abstractmethod

class IResource(ABC):

    @abstractmethod
    def close(self):
        pass