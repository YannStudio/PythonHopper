"""Unit tests for the BaseDB class and database abstraction pattern."""

import json
import os
import tempfile
from dataclasses import dataclass, asdict
from typing import Any, Dict, List

import pytest

from database.base import BaseDB


@dataclass
class SampleItem:
    """Simple test data class."""
    name: str
    value: int
    tags: List[str] = None
    
    def __post_init__(self):
        if self.tags is None:
            self.tags = []
    
    @classmethod
    def from_any(cls, data: Dict[str, Any]) -> "SampleItem":
        """Create from dictionary."""
        if isinstance(data, cls):
            return data
        return cls(
            name=str(data.get("name", "")),
            value=int(data.get("value", 0)),
            tags=data.get("tags", [])
        )


class SampleDB(BaseDB[SampleItem]):
    """Concrete implementation for testing."""
    
    def schema_version(self) -> str:
        return "1.0"
    
    def to_dict(self, item: SampleItem) -> Dict[str, Any]:
        d = asdict(item)
        return d
    
    def from_dict(self, d: Dict[str, Any]) -> SampleItem:
        return SampleItem.from_any(d)


@pytest.fixture
def temp_db_file():
    """Create a temporary database file."""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        temp_path = f.name
    yield temp_path
    # Cleanup
    if os.path.exists(temp_path):
        os.unlink(temp_path)


class TestBaseDBInitialization:
    """Test BaseDB initialization."""
    
    def test_empty_initialization(self):
        """Test creating empty database."""
        db = SampleDB()
        assert db.items == []
    
    def test_initialization_with_items(self):
        """Test creating database with initial items."""
        items = [SampleItem("item1", 10), SampleItem("item2", 20)]
        db = SampleDB(items)
        assert len(db.items) == 2
        assert db.items[0].name == "item1"


class TestBaseDBPersistence:
    """Test loading and saving functionality."""
    
    def test_empty_db_save_and_load(self, temp_db_file):
        """Test saving and loading empty database."""
        db = SampleDB()
        db.save(temp_db_file)
        
        # Verify file exists and has correct structure
        assert os.path.exists(temp_db_file)
        with open(temp_db_file) as f:
            data = json.load(f)
        assert "schema" in data
        assert "items" in data
        assert data["items"] == []
        
        # Load and verify
        db2 = SampleDB()
        db2.load(temp_db_file)
        assert db2.items == []
    
    def test_save_and_load_with_items(self, temp_db_file):
        """Test saving and loading database with items."""
        original_items = [
            SampleItem("Alice", 100, ["tag1", "tag2"]),
            SampleItem("Bob", 200, ["tag2", "tag3"]),
        ]
        db = SampleDB(original_items)
        db.save(temp_db_file)
        
        # Load into new database
        db2 = SampleDB()
        db2.load(temp_db_file)
        
        assert len(db2.items) == 2
        assert db2.items[0].name == "Alice"
        assert db2.items[0].value == 100
        assert db2.items[0].tags == ["tag1", "tag2"]
        assert db2.items[1].name == "Bob"
        assert db2.items[1].value == 200
    
    def test_load_nonexistent_file(self):
        """Test loading from non-existent file gracefully."""
        db = SampleDB()
        db.load("/nonexistent/path/to/database.json")
        assert db.items == []
    
    def test_corrupted_json_handling(self, temp_db_file):
        """Test handling of corrupted JSON file."""
        # Write invalid JSON
        with open(temp_db_file, 'w') as f:
            f.write("{ invalid json ]")
        
        db = SampleDB()
        db.load(temp_db_file)
        assert db.items == []
    
    def test_old_format_compatibility(self, temp_db_file):
        """Test loading old flat list JSON format."""
        # Write old format (just a list)
        old_data = [
            {"name": "Old1", "value": 10, "tags": []},
            {"name": "Old2", "value": 20, "tags": ["old"]},
        ]
        with open(temp_db_file, 'w') as f:
            json.dump(old_data, f)
        
        # Load should handle old format
        db = SampleDB()
        db.load(temp_db_file)
        assert len(db.items) == 2
        assert db.items[0].name == "Old1"
        assert db.items[1].name == "Old2"
    
    def test_schema_version_in_output(self, temp_db_file):
        """Test that schema version is written to file."""
        db = SampleDB([SampleItem("test", 42)])
        db.save(temp_db_file)
        
        with open(temp_db_file) as f:
            data = json.load(f)
        assert data["schema"] == "1.0"
    
    def test_unicode_handling(self, temp_db_file):
        """Test that Unicode characters are preserved."""
        items = [
            SampleItem("café", 1, ["français"]),
            SampleItem("日本", 2, ["日本語"]),
            SampleItem("Москва", 3, ["русский"]),
        ]
        db = SampleDB(items)
        db.save(temp_db_file)
        
        db2 = SampleDB()
        db2.load(temp_db_file)
        assert db2.items[0].name == "café"
        assert db2.items[1].name == "日本"
        assert db2.items[2].name == "Москва"
    
    def test_directory_creation(self):
        """Test that parent directories are created on save."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create a path with non-existent parents
            nested_path = os.path.join(tmpdir, "a", "b", "c", "database.json")
            db = SampleDB([SampleItem("test", 1)])
            
            db.save(nested_path)
            assert os.path.exists(nested_path)
            
            # Verify we can load it back
            db2 = SampleDB()
            db2.load(nested_path)
            assert len(db2.items) == 1


class TestBaseDBDataIntegrity:
    """Test data integrity and round-trip conversion."""
    
    def test_roundtrip_conversion(self):
        """Test converting item → dict → item preserves data."""
        original = SampleItem("roundtrip", 999, ["a", "b", "c"])
        db = SampleDB()
        
        # Convert to dict
        item_dict = db.to_dict(original)
        
        # Convert back
        restored = db.from_dict(item_dict)
        
        assert restored.name == original.name
        assert restored.value == original.value
        assert restored.tags == original.tags
    
    def test_handling_missing_optional_fields(self):
        """Test handling of missing optional fields in loaded data."""
        db = SampleDB()
        
        # Dict missing 'tags' field (optional)
        incomplete_dict = {"name": "incomplete", "value": 5}
        item = db.from_dict(incomplete_dict)
        
        assert item.name == "incomplete"
        assert item.value == 5
        assert item.tags == []


class TestTypeGenericSupport:
    """Test that the generic type system works correctly."""
    
    def test_database_maintains_type_information(self):
        """Test that items maintain their type through operations."""
        items = [SampleItem("test", 1)]
        db = SampleDB(items)
        
        retrieved = db.items[0]
        assert isinstance(retrieved, SampleItem)
        assert hasattr(retrieved, 'name')
        assert hasattr(retrieved, 'value')
        assert hasattr(retrieved, 'tags')


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
