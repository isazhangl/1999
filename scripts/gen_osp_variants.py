#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gera os variantes (Formato × Quilates × Metal) de cada engaste da linha
O SIM PERFEITO, precificados pela fórmula em docs/定价公式-O-SIM-PERFEITO.md.
Saída: um JSON por engaste em <outdir>/<CODE>.json com {title, productType, tags, options, variants}.
"""
import json, os, sys

# --- parâmetros (espelham o doc de preços) ---
PEDRA = {"050":1400, "070":2000, "100":3200, "150":5400, "200":8000, "250":11500, "300":15500}
FATOR_METAL = {"AG":0.55, "OB":1.00, "OA":1.00, "OR":1.00, "PL":1.35}
PISO = 1990

BASE = {
    "SOL":1600,"ENL":1900,"BZL":1900,"CAT":1900,"INF":2000,
    "TRG":3200,"TRE":3200,"TRB":3200,"TRR":3200,
    "HAL":2600,"HAF":2800,"HAP":3000,"HAV":3000,
    "HHL":2600,"HHP":3000,"HHC":3000,
    "PAV":2400,"VIN":2600,"CRO":2600,
}

# nome PT de cada engaste (para título do produto)
NOME = {
    "SOL":"Solitário Puro","ENL":"Enlace","BZL":"Aura","CAT":"Catedral","INF":"Infinito",
    "TRG":"Trilogia Gota","TRE":"Trilogia Esmeralda","TRB":"Trilogia Baguete","TRR":"Trilogia Redonda",
    "HAL":"Halo","HAF":"Halo Flor","HAP":"Halo Pavé","HAV":"Halo Vinha",
    "HHL":"Hidden Halo","HHP":"Halo Oculto Pavé","HHC":"Halo Oculto Catedral",
    "PAV":"Pavé","VIN":"Vinha","CRO":"Coroa",
}

# formatos: code -> (nome PT)
FORMATOS = [
    ("RD","Redondo"),("OV","Oval"),("EM","Esmeralda"),("MQ","Marquise"),("GO","Gota"),
    ("PR","Princesa"),("RA","Radiante"),("AL","Almofada"),("AS","Asscher"),("CR","Coração"),
]
CARATS = [("050","0,5 ct"),("070","0,7 ct"),("100","1,0 ct"),("150","1,5 ct"),
          ("200","2,0 ct"),("250","2,5 ct"),("300","3,0 ct")]
METAIS = [("AG","Prata 925"),("OB","Ouro 18K branco"),("OA","Ouro 18K amarelo"),
          ("OR","Ouro 18K rosé"),("PL","Platina 950")]

def preco(code, car, met):
    bruto = (BASE[code] + PEDRA[car]) * FATOR_METAL[met]
    arred = round(bruto / 100.0) * 100 - 10     # nearest 100, -10 -> termina em 90
    return max(arred, PISO)

def build(code):
    variants = []
    for fc, fn in FORMATOS:
        for cc, cn in CARATS:
            for mc, mn in METAIS:
                sku = f"SP-AN-{code}-{fc}-{cc}-{mc}"
                variants.append({
                    "sku": sku,
                    "price": f"{preco(code, cc, mc)}.00",
                    "optionValues": [
                        {"optionName":"Formato","name":fn},
                        {"optionName":"Quilates","name":cn},
                        {"optionName":"Metal","name":mn},
                    ],
                })
    return {
        "title": f"Anel {NOME[code]}",
        "productType": "Anel de Noivado",
        "vendor": "1999",
        "tags": ["o-sim-perfeito","anel-de-noivado","diamante-cultivado", f"engaste-{code.lower()}"],
        "options": ["Formato","Quilates","Metal"],
        "variants": variants,
    }

if __name__ == "__main__":
    outdir = sys.argv[1] if len(sys.argv) > 1 else "."
    os.makedirs(outdir, exist_ok=True)
    total = 0
    for code in BASE:
        data = build(code)
        total += len(data["variants"])
        with open(os.path.join(outdir, f"{code}.json"), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        print(f"{code}: {len(data['variants'])} variantes  preço "
              f"{data['variants'][0]['price']} .. {data['variants'][-1]['price']}")
    print(f"TOTAL: {len(BASE)} engastes, {total} variantes")
