"""Test supplier filtering functionality."""
import pytest
from models import Supplier
from suppliers_db import SuppliersDB


def test_supplier_find_product_type_filter():
    """Test filtering by product type."""
    suppliers = [
        Supplier(supplier="Supplier A", product_type="Staal", product_description="Profiel"),
        Supplier(supplier="Supplier B", product_type="Aluminium", product_description="Plaat"),
        Supplier(supplier="Supplier C", product_type="Staal", product_description="Buis"),
    ]
    db = SuppliersDB(suppliers)
    
    # Filter by product type "Staal"
    results = db.find("", product_type_filter="Staal")
    assert len(results) == 2
    assert results[0].supplier == "Supplier A"
    assert results[1].supplier == "Supplier C"


def test_supplier_find_description_filter():
    """Test filtering by product description."""
    suppliers = [
        Supplier(supplier="Supplier A", product_type="Staal", product_description="Profiel"),
        Supplier(supplier="Supplier B", product_type="Aluminium", product_description="Plaat"),
        Supplier(supplier="Supplier C", product_type="Staal", product_description="Profiel"),
    ]
    db = SuppliersDB(suppliers)
    
    # Filter by description "Profiel"
    results = db.find("", product_desc_filter="Profiel")
    assert len(results) == 2
    assert results[0].supplier == "Supplier A"
    assert results[1].supplier == "Supplier C"


def test_supplier_find_combined_filters_and_logic():
    """Test AND logic: both product type AND description must match."""
    suppliers = [
        Supplier(supplier="Supplier A", product_type="Staal", product_description="Profiel"),
        Supplier(supplier="Supplier B", product_type="Aluminium", product_description="Profiel"),
        Supplier(supplier="Supplier C", product_type="Staal", product_description="Buis"),
    ]
    db = SuppliersDB(suppliers)
    
    # Filter by product type "Staal" AND description "Profiel"
    results = db.find("", product_type_filter="Staal", product_desc_filter="Profiel")
    assert len(results) == 1
    assert results[0].supplier == "Supplier A"


def test_supplier_find_combined_filters_no_matches():
    """Test AND logic returns empty when combination doesn't exist."""
    suppliers = [
        Supplier(supplier="Supplier A", product_type="Staal", product_description="Profiel"),
        Supplier(supplier="Supplier B", product_type="Aluminium", product_description="Profiel"),
        Supplier(supplier="Supplier C", product_type="Staal", product_description="Buis"),
    ]
    db = SuppliersDB(suppliers)
    
    # Filter by product type "Aluminium" AND description "Buis" (doesn't exist)
    results = db.find("", product_type_filter="Aluminium", product_desc_filter="Buis")
    assert len(results) == 0


def test_supplier_find_case_insensitive():
    """Test that filters are case-insensitive."""
    suppliers = [
        Supplier(supplier="Supplier A", product_type="Staal", product_description="Profiel"),
        Supplier(supplier="Supplier B", product_type="Aluminium", product_description="Plaat"),
    ]
    db = SuppliersDB(suppliers)
    
    # Different cases should still match
    results = db.find("", product_type_filter="STAAL", product_desc_filter="profiel")
    assert len(results) == 1
    assert results[0].supplier == "Supplier A"


def test_supplier_find_with_query_and_filters():
    """Test combining text search with product filters."""
    suppliers = [
        Supplier(supplier="Steel Corp", product_type="Staal", product_description="Profiel"),
        Supplier(supplier="Steel Plus", product_type="Staal", product_description="Buis"),
        Supplier(supplier="Aluminium Works", product_type="Aluminium", product_description="Profiel"),
    ]
    db = SuppliersDB(suppliers)
    
    # Search for "Steel" with product type filter "Staal"
    results = db.find("Steel", product_type_filter="Staal")
    assert len(results) == 2
    assert results[0].supplier == "Steel Corp"
    assert results[1].supplier == "Steel Plus"
    
    # Search for "Steel" with product type "Staal" AND description "Buis"
    results = db.find("Steel", product_type_filter="Staal", product_desc_filter="Buis")
    assert len(results) == 1
    assert results[0].supplier == "Steel Plus"


def test_supplier_find_empty_filters():
    """Test that empty filters don't restrict results."""
    suppliers = [
        Supplier(supplier="Supplier A", product_type="Staal", product_description="Profiel"),
        Supplier(supplier="Supplier B", product_type="Aluminium", product_description="Plaat"),
    ]
    db = SuppliersDB(suppliers)
    
    # Empty filters should return all suppliers
    results = db.find("", product_type_filter="", product_desc_filter="")
    assert len(results) == 2


def test_supplier_find_favorites_sorted_first():
    """Test that favorite suppliers are sorted first."""
    suppliers = [
        Supplier(supplier="Supplier B", product_type="Staal", product_description="Profiel", favorite=False),
        Supplier(supplier="Supplier A", product_type="Staal", product_description="Profiel", favorite=True),
    ]
    db = SuppliersDB(suppliers)
    
    results = db.find("", product_type_filter="Staal")
    assert results[0].supplier == "Supplier A"  # Favorite first
    assert results[1].supplier == "Supplier B"


def test_get_product_descriptions_for_type():
    """Test getting descriptions filtered by product type."""
    suppliers = [
        Supplier(supplier="Supplier A", product_type="Staal", product_description="Profiel"),
        Supplier(supplier="Supplier B", product_type="Staal", product_description="Buis"),
        Supplier(supplier="Supplier C", product_type="Aluminium", product_description="Profiel"),
    ]
    db = SuppliersDB(suppliers)
    
    # Get descriptions for "Staal"
    descs = db.get_product_descriptions_for_type("Staal")
    assert len(descs) == 2
    assert "Profiel" in descs
    assert "Buis" in descs
    assert "Profiel" in descs  # Profiel appears in both types but should only appear once here
    
    # Get descriptions for "Aluminium"
    descs = db.get_product_descriptions_for_type("Aluminium")
    assert len(descs) == 1
    assert "Profiel" in descs


def test_get_product_descriptions_for_empty_type():
    """Test that empty type returns all descriptions."""
    suppliers = [
        Supplier(supplier="Supplier A", product_type="Staal", product_description="Profiel"),
        Supplier(supplier="Supplier B", product_type="Aluminium", product_description="Buis"),
    ]
    db = SuppliersDB(suppliers)
    
    # Empty type should return all descriptions
    descs = db.get_product_descriptions_for_type("")
    assert len(descs) == 2
    assert "Profiel" in descs
    assert "Buis" in descs
