from dataclasses import dataclass
from typing import Generic, TypeVar, Optional

T = TypeVar("T")


@dataclass
class Result(Generic[T]):
    @dataclass
    class Error:
        msg: str

    data: T
    error: Optional[Error] = None

    def __json__(self):
        return {
            "data": self.data,
            "error": self.error.msg if self.error is not None else None,
        }
