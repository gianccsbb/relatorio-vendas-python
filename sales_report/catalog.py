"""Product catalog: canonical names, categories and known spelling variants."""

CATALOG = {
    # canonical name: (category, list price in cents)
    "Café coado": ("Cafés", 600),
    "Espresso": ("Cafés", 700),
    "Cappuccino": ("Cafés", 1200),
    "Café gelado com rapadura": ("Cafés", 1400),
    "Pão na chapa": ("Padaria", 700),
    "Tapioca de queijo coalho": ("Padaria", 1400),
    "Cuscuz com ovo": ("Padaria", 1300),
    "Croissant": ("Padaria", 1100),
    "Bolo de rolo (fatia)": ("Doces", 1000),
    "Bolo Souza Leão": ("Doces", 1200),
    "Cartola": ("Doces", 1600),
    "Brigadeiro": ("Doces", 400),
}

# Spellings seen in the POS and delivery exports that fuzzy matching alone would not resolve.
# Keys are normalized (lowercase, no accents, single spaces).
ALIASES = {
    "expresso": "Espresso",
    "capuccino": "Cappuccino",
    "capucino": "Cappuccino",
    "cafe gelado": "Café gelado com rapadura",
    "gelado rapadura": "Café gelado com rapadura",
    "pao chapa": "Pão na chapa",
    "pao na chapa c/ manteiga": "Pão na chapa",
    "tapioca queijo coalho": "Tapioca de queijo coalho",
    "tapioca coalho": "Tapioca de queijo coalho",
    "cuscuz c/ ovo": "Cuscuz com ovo",
    "bolo de rolo": "Bolo de rolo (fatia)",
    "bolo de rolo fatia": "Bolo de rolo (fatia)",
    "bolo souza leao": "Bolo Souza Leão",
    "souza leao": "Bolo Souza Leão",
    "brigadeiro un": "Brigadeiro",
    "brigadeiro unid": "Brigadeiro",
}

PAYMENT_ALIASES = {
    "pix": "PIX",
    "dinheiro": "Dinheiro",
    "especie": "Dinheiro",
    "cartao debito": "Débito",
    "debito": "Débito",
    "cartao credito": "Crédito",
    "credito": "Crédito",
    "app": "App de entrega",
    "app de entrega": "App de entrega",
}
