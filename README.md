# Retail Fulfillment Warehouse

Mobile-first Flask + PostgreSQL application for retail warehouse fulfillment.

## Render
This repository includes `render.yaml`. In Render, create a Blueprint from this repository. It creates:
- Free Python web service
- Free Render Postgres database

The Free Postgres database is intended for testing and expires after 30 days. Upgrade the database before expiry for long-term use.

Build: `pip install -r requirements.txt`
Start: `gunicorn app:app`

Health: `/health`

## Excel
Barang headers: SKU, Nama Barang, Kategori, Satuan, Harga Pembelian, Supplier, Stok Minimum
Request headers: SKU, Nama Barang, Qty Request
