"""Keep test artifacts inside the workspace and preserve Windows parent ACLs."""
import shutil
import uuid
from pathlib import Path


class TemporaryWorkspace:
    def __init__(self, prefix='test-'):
        self.parent = (Path(__file__).resolve().parents[1] / '.test-output').resolve()
        self.parent.mkdir(exist_ok=True)
        self.path = self.parent / (prefix + uuid.uuid4().hex)
        self.path.mkdir()
        self.name = str(self.path)

    def cleanup(self):
        # Delete only the unique directory this helper created, never its parent.
        resolved = self.path.resolve()
        if resolved.parent != self.parent or resolved == self.parent:
            raise RuntimeError('Refusing to clean a test directory outside .test-output')
        shutil.rmtree(resolved)
