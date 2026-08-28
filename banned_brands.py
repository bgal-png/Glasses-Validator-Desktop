"""Per-site banned / allowed brand lists, ported verbatim from the web validator."""
from __future__ import annotations

SITE_BANNED = {
    # ── Czech sites ──
    "Čočky-kontaktni.cz": {
        "type": "banned",
        "brands": {"Calvin Klein", "Dolce & Gabbana", "Chiara Ferragni", "Jimmy Choo", "Lacoste", "Marisio", "Missoni", "Montblanc", "Persol", "Meller", "Celine"},
    },
    "Čočky-online.cz": {
        "type": "banned",
        "brands": {"Gucci", "Chiara Ferragni", "Christian Dior", "Julbo", "Just Cavalli", "Montblanc", "Meller", "Celine"},
    },
    "Čočky-optika.cz": {
        "type": "banned",
        "brands": {"Givenchy", "Havaianas", "Christian Dior", "Julbo", "Just Cavalli", "Kate Spade", "Meller", "Celine"},
    },
    "Alensa.cz": {
        "type": "banned",
        "brands": {"Meller", "Celine"},
    },
    "Kontaktni.cz": {
        "type": "banned",
        "brands": {"Meller", "Celine"},
    },
    # ── Poland ──
    "Alensa.pl": {
        "type": "banned",
        "brands": {"Hawkers", "Meller", "Celine"},
    },
    # ── Greece ──
    "Alensa.gr": {
        "type": "banned",
        "brands": {"Hawkers", "Meller", "Celine"},
    },
    "Mataki.gr": {
        "type": "banned",
        "brands": {"Hawkers", "Meller", "Celine"},
    },
    # ── France ──
    "Alensa.fr": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    # ── Spain ──
    "Alensa.es": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    "Lentes-de-contacto.es": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    "Lentes-shop.es": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    # ── Italy ──
    "Alensa.it": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    "Adrialenti.it": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    "Lenti-ottica.it": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    # ── Croatia ──
    "Alensa.hr": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    "Adrialece.hr": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    # ── Slovenia ──
    "Alensa.si": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    "Moje-lece.si": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    # ── Serbia ──
    "Alensa.rs": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    # ── Bosnia ──
    "Adrialece.ba": {
        "type": "banned",
        "brands": {"Hawkers", "Celine"},
    },
    # ── Portugal ──
    "Alensa.pt": {
        "type": "banned",
        "brands": {"Meller", "Celine"},
    },
    # ── Romania ──
    "Contact-lentile.ro": {
        "type": "banned",
        "brands": {"Ana Hickman", "Morel", "Celine"},
    },
    "Videt.ro": {
        "type": "banned",
        "brands": {"Ana Hickman", "Morel", "Celine"},
    },
    "Xlentile.ro": {
        "type": "banned",
        "brands": {"Ana Hickman", "Morel", "Celine"},
    },
    # ── Madagascar ──
    "vallis.mg": {
        "type": "banned",
        "brands": {"Desiree", "Celine"},
    },
    # ── International ──
    "Adrial.eu": {
        "type": "banned",
        "brands": {"Desiree", "Celine"},
    },
    "Alensa.com": {
        "type": "banned",
        "brands": {"Hawkers", "Meller", "Celine"},
    },
    # ── Norway ──
    "Alensa.no": {
        "type": "banned",
        "brands": {
            "Adidas", "Alexander McQueen", "Balenciaga", "Beron", "Boss by Hugo Boss",
            "Burberry", "Calvin Klein", "Carolina Herrera", "Carrera", "Celine",
            "Chloe", "Christian Dior", "David Beckham", "Dolce & Gabbana", "Dsquared2",
            "Giorgio Armani", "Gucci", "Hugo by Hugo Boss", "Jimmy Choo", "Kate Spade",
            "Love Moschino", "Marc Jacobs", "Maui Jim", "Max Mara", "Missoni",
            "Miu Miu", "Montblanc", "Moschino", "Oakley", "Persol",
            "Police", "Polo Ralph Lauren", "Prada", "Prada Linea Rossa", "Ralph Lauren",
            "Ray-Ban", "Saint Laurent", "Serengeti", "Swarovski", "Tiffany & Co.",
            "Tom Ford", "Tommy Hilfiger", "Versace", "Victoria Beckham",
        },
    },
    # ── Ukraine ──
    "Alensa.ua": {
        "type": "allowed",
        "brands": {"Crullé", "Marisio", "Kimikado", "Lewish", "Beron", "Válle", "Polaroid"},
    },
}
