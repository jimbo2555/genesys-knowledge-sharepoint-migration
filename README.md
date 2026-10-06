# Genesys Knowledge to SharePoint Migration

Proof of concept converter for migrating Genesys Cloud Knowledge v2 content to Microsoft SharePoint.

The converter reads a Genesys Knowledge JSON export and creates individual DOCX files.

## Features

* Converts individual Genesys knowledge documents to DOCX
* Preserves text formatting
* Preserves headings and paragraphs
* Converts lists and tables
* Preserves hyperlinks
* Downloads externally referenced images
* Embeds images directly into DOCX
* Generates a migration report
* Retains Genesys metadata for subsequent SharePoint mapping

## Installation

python3 -m pip install -r requirements.txt

## Usage

python3 genesys_kb_to_docx.py "export.json" -o "sharepoint_docs"

## Status

Proof of concept. Further development is required before production migration use.
