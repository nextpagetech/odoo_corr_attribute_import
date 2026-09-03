{
    "name": "Corrugation Attribute Import",
    "summary": "Import product attributes and values from XLSX/CSV without technical IDs",
    "description": """
Imports product attributes and attribute values from XLSX or CSV files using names instead of technical IDs, with options to update existing records and create missing attributes.
    """,
    "version": "18.0.1.0.0",
    "author": "Next Page Technologies Pvt Ltd",
    "website": "https://nextpagetechnologies.com",
    "license": "LGPL-3",
    "category": "Productivity",
    "depends": ["product"],
    "external_dependencies": {"python": ["openpyxl"]},
    "data": [
        "security/ir.model.access.csv",
        "views/attribute_import_views.xml",
    ],
    "installable": True,
    "application": True,
}
