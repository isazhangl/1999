#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Constrói o JSONL para bulkOperationRunMutation(productSet) — 1 linha por engaste.
Cada linha: {"input": ProductSetInput} com productOptions + 350 variants.
Uso: python3 osp_bulk_jsonl.py <saida.jsonl>
"""
import json, sys
from gen_osp_variants import build, BASE, FORMATOS, CARATS, METAIS

def product_options():
    return [
        {"name":"Formato","values":[{"name":n} for _,n in FORMATOS]},
        {"name":"Quilates","values":[{"name":n} for _,n in CARATS]},
        {"name":"Metal","values":[{"name":n} for _,n in METAIS]},
    ]

def main(out):
    with open(out, "w", encoding="utf-8") as f:
        for code in BASE:
            d = build(code)
            inp = {
                "title": d["title"],
                "status": "DRAFT",
                "vendor": d["vendor"],
                "productType": d["productType"],
                "tags": d["tags"],
                "productOptions": product_options(),
                "variants": d["variants"],
            }
            f.write(json.dumps({"input": inp}, ensure_ascii=False) + "\n")
    print("wrote", out)

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "osp_bulk.jsonl")
