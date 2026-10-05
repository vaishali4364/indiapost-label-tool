def get_item_specs(title):
  t = str(title).lower()

  # Detect pack multiplier from title (e.g. "Pack of 2", "(2)", "2x", etc.)
  pack_mult = 1
  m_pack = re.search(r"pack\s*(?:of)?\s*(\d+)", t)
  m_lead = re.match(r"^(\d+)\s+", t)
  m_paren = re.search(r"[\(\[]\s*(\d+)\s*[\)\]]", t)
  m_x = re.search(r"(\d+)\s*x\s*(?!20ml|100ml)", t)

  if m_pack:
    pack_mult = int(m_pack.group(1))
  elif m_lead and int(m_lead.group(1)) in [2, 3, 4, 5, 6, 8]:
    pack_mult = int(m_lead.group(1))
  elif m_paren and int(m_paren.group(1)) in [2, 3, 4, 5, 6, 8]:
    pack_mult = int(m_paren.group(1))
  elif m_x and int(m_x.group(1)) in [2, 3, 4, 5, 6, 8]:
    pack_mult = int(m_x.group(1))

  # --- Olivia Herbal Bleach Cream ---
  if "olivia" in t and "bleach" in t:
    return {
        "desc": "Olivia Herbal Bleach Cream",
        "hs": "33049910",
        "wt": 49 * pack_mult,
        "val": 55 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Betadine Antiseptic ---
  if "betadine" in t:
    return {
        "desc": "Betadine Povidone-iodine",
        "hs": "30049099",
        "wt": 248 if pack_mult == 2 else 140 * pack_mult,
        "val": 90 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Fragrances & Perfumes (Bellavita, Wild Stone, Fogg) ---
  if any(k in t for k in ["bella vita", "bellavita", "perfume", "oud", "edp"]):
    hs = "33030040"
    if "intense" in t and ("ceo" in t or "100" in t):
      return {
          "desc": "Bellavita CEO Intense 100ml",
          "hs": hs,
          "wt": 448 * pack_mult,
          "val": 450 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    if any(
        k in t
        for k in ["combo", "2 pack", "2 x 100ml", "ceo + goat", "ceo and goat"]
    ):
      return {
          "desc": "Bellavita CEO & GOAT 2Pk",
          "hs": hs,
          "wt": 848,
          "val": 700,
          "is_multi": True,
      }
    if any(k in t for k in ["gift set", "4 x 20ml", "4x20ml", "all star"]):
      if (
          pack_mult == 2
          or "pack of 2" in t
          or "2 bella vita" in t
          or "set 2" in t
      ):
        return {
            "desc": "Bellavita Gift Set 4x20ml 2Pk",
            "hs": hs,
            "wt": 848,
            "val": 700,
            "is_multi": True,
        }
      return {
          "desc": "Bellavita Gift Set 4x20ml",
          "hs": hs,
          "wt": 398,
          "val": 350,
          "is_multi": False,
      }
    if any(
        k in t
        for k in [
            "2 x 20ml",
            "2x20ml",
            "white and honey",
            "white oud & honey",
            "fresh + white",
            "date and senorita",
            "ceo + white",
            "date & glam",
        ]
    ):
      return {
          "desc": "Bellavita Perfume Duo 2x20ml",
          "hs": hs,
          "wt": 198,
          "val": 360,
          "is_multi": True,
      }
    if "20ml" in t:
      return {
          "desc": "Bellavita Pocket Perfume 20ml",
          "hs": hs,
          "wt": 98 * pack_mult,
          "val": 180 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    if any(k in t for k in ["dark oud", "oud dark", "oud gold"]):
      return {
          "desc": "Bellavita Dark Oud EDP 100ml",
          "hs": hs,
          "wt": 498 * pack_mult,
          "val": 486 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    return {
        "desc": "Bellavita Eau De Parfum 100ml",
        "hs": hs,
        "wt": 398 * pack_mult,
        "val": 350 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Skincare, Serums & Face Packs ---
  if any(k in t for k in ["detan", "moroccon", "face pack", "sayy"]):
    return {
        "desc": "Sayy Moroccan DeTan Mask",
        "hs": "33049090",
        "wt": 140 * pack_mult,
        "val": 210 * pack_mult,
        "is_multi": pack_mult > 1,
    }
  if any(k in t for k in ["elvive", "serum", "loreal paris", "biolage"]):
    if "combo" in t:
      return {
          "desc": "Loreal Elvive Hair Combo",
          "hs": "33059090",
          "wt": 548,
          "val": 506,
          "is_multi": True,
      }
    if pack_mult > 1:
      wt_map = {2: 240, 3: 348, 4: 490}
      return {
          "desc": "Loreal Paris Elvive Serum",
          "hs": "33059090",
          "wt": wt_map.get(pack_mult, 140 * pack_mult),
          "val": 434 * pack_mult,
          "is_multi": True,
      }
    return {
        "desc": "Loreal Paris Elvive Serum",
        "hs": "33059090",
        "wt": 140,
        "val": 434,
        "is_multi": False,
    }
  if any(k in t for k in ["gluta hya", "gluta-hya", "vaseline"]):
    if any(k in t for k in ["bundle", "pack of 3", "(3)"]) or pack_mult == 3:
      return {
          "desc": "Vaseline Gluta Hya 3Pk",
          "hs": "33049090",
          "wt": 645,
          "val": 826,
          "is_multi": True,
      }
    if "400ml" in t or "sun protect" in t:
      return {
          "desc": "Vaseline Body Lotion 400ml",
          "hs": "33049090",
          "wt": 480 * pack_mult,
          "val": 340 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    if pack_mult > 1:
      return {
          "desc": "Vaseline Gluta Hya Lotion",
          "hs": "33049090",
          "wt": 398 if pack_mult == 2 else 249 * pack_mult,
          "val": 255 * pack_mult,
          "is_multi": True,
      }
    return {
        "desc": "Vaseline Gluta Hya Lotion",
        "hs": "33049090",
        "wt": 249,
        "val": 255,
        "is_multi": False,
    }
  if any(k in t for k in ["roll on", "roll-on", "chemist at play", "rexona"]):
    return {
        "desc": "Deodorant Underarm Roll On",
        "hs": "33072000",
        "wt": 98 * pack_mult,
        "val": 291 * pack_mult,
        "is_multi": pack_mult > 1,
    }
  if "supradyn" in t:
    return {
        "desc": "Supradyn Multivitamin 60s",
        "hs": "21069099",
        "wt": 98 * pack_mult,
        "val": 250 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Shoe Polish & Sponges ---
  if "sponge" in t:
    return {
        "desc": "Kiwi Shoe Shine Sponge",
        "hs": "34051000",
        "wt": 198 if pack_mult == 2 else 98 * pack_mult,
        "val": 170 * pack_mult,
        "is_multi": pack_mult > 1,
    }
  if any(k in t for k in ["kiwi", "cherry blossom", "shoe polish"]):
    if any(k in t for k in ["instant", "liquid"]):
      return {
          "desc": "Kiwi Instant Shoe Polish",
          "hs": "34051000",
          "wt": 198 if pack_mult == 2 else 98 * pack_mult,
          "val": 106 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    return {
        "desc": "Kiwi Shoe Polish Wax Tin",
        "hs": "34051000",
        "wt": 248 if pack_mult == 2 else 120 * pack_mult,
        "val": 106 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Copperware ---
  if "copper" in t:
    if any(k in t for k in ["pot", "jug"]):
      return {
          "desc": "Pure Copper Water Pot",
          "hs": "74198090",
          "wt": 326,
          "val": 945,
          "is_multi": False,
      }
    wt_balls = 49 if pack_mult <= 2 else (98 if pack_mult <= 4 else 198)
    return {
        "desc": "Ayurvedic Copper Ball",
        "hs": "74198090",
        "wt": wt_balls,
        "val": 162 * (pack_mult // 2 if pack_mult > 1 else 1),
        "is_multi": pack_mult > 1,
    }

  # --- Shaving Razors & Blades ---
  if any(k in t for k in ["guard", "mach 3", "fusion", "vector", "razor"]):
    return {
        "desc": "Gillette Shaving Razor",
        "hs": "82121010",
        "wt": 198 if pack_mult == 2 else 98 * pack_mult,
        "val": 260 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Oral Care ---
  if any(k in t for k in ["toothbrush", "oral b", "oral-b"]):
    return {
        "desc": "Oral Care Toothbrush Pack",
        "hs": "96032100",
        "wt": 198 if pack_mult == 2 else 98 * pack_mult,
        "val": 213 * pack_mult,
        "is_multi": pack_mult > 1,
    }
  if any(k in t for k in ["paste", "dabur", "colgate"]):
    return {
        "desc": "Dental Toothpaste Tube",
        "hs": "33061020",
        "wt": 198 * pack_mult,
        "val": 210 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # Default fallback
  return {
      "desc": str(title)[:28],
      "hs": "33049990",
      "wt": 140 * pack_mult,
      "val": 250 * pack_mult,
      "is_multi": pack_mult > 1,
  }