from typing import override
from .i_resource import IResource


class NullResource(IResource):

    @override
    def close(self):
        print("resource exception")