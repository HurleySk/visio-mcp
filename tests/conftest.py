import pytest

from tests.fakes import FakeConnector
from visio_mcp import com_runtime


@pytest.fixture
def connector():
    fake = FakeConnector()
    com_runtime.configure(lambda: fake, sleep=lambda seconds: None)
    yield fake
    com_runtime.reset()
