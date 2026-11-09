"""Tests of the cart pattern and of the page name in the report.

python -m unittest test_cart_label      (run with the Python that has Pillow, reportlab and playwright)
A page called "cart" in the report must be a cart by its address; every other page is named from its title or, with
no title, from the last part of its address. The negative cases are real addresses where "cart" is the start of a word.
"""
import unittest

import quick as Q

NOT_CART = [
    ("https://www.leslipfrancais.fr/collections/carte-cadeau", "Carte cadeau \u2013 Le Slip Fran\u00e7ais", "Carte cadeau"),
    ("https://www.balibaris.com/products/e-carte-cadeau-1", "E-carte cadeau | Balibaris", "E-carte cadeau"),
    ("https://www.ldlc.com/informatique/pieces-informatique/carte-graphique-interne/c4684/", "", "c4684"),
    ("https://probikeshop.com/collections/vtt-cartouches-et-pieces", "VTT cartouches et pi\u00e8ces - Probikeshop",
     "VTT cartouches et pi\u00e8ces"),
]
CART = ["https://bonnegueule.fr/cart", "https://deguisetoi.fr/cart/", "https://www.easons.com/checkout/basket",
        "https://www.armandthiery.fr/fr/panier.cfm", "https://x.nl/winkelwagen", "https://x.com/shopping-cart"]


class CartPattern(unittest.TestCase):
    def test_real_carts_match(self):
        for url in CART:
            self.assertTrue(Q.CART_HINT.search(url), url)

    def test_words_starting_with_cart_do_not_match(self):
        for url, _, _ in NOT_CART:
            self.assertFalse(Q.CART_HINT.search(url), url)


class PageLabel(unittest.TestCase):
    def test_cart_by_address_keeps_the_name(self):
        for url in CART:
            self.assertEqual(Q.page_label({"type": "cart", "url": url, "title": "Panier"}), "cart", url)

    def test_scanned_as_cart_but_not_a_cart_is_named_from_the_page(self):
        for url, title, name in NOT_CART:
            label = Q.page_label({"type": "cart", "url": url, "title": title})
            self.assertEqual(label, f"page \u201c{name}\u201d", url)
            self.assertNotIn("cart", label.lower().replace("carte", "").replace("cartouches", "").replace("cart.", ""))
            self.assertNotEqual(label, "cart")

    def test_other_types_keep_their_names(self):
        self.assertEqual(Q.page_label({"type": "home", "url": "https://x.fr/", "title": "x"}), "home page")
        self.assertEqual(Q.page_label({"type": "contact", "url": "https://x.fr/contact"}), "contact page")

    def test_final_address_decides(self):
        p = {"type": "cart", "url": "https://x.fr/cart", "finalUrl": "https://x.fr/collections/carte-cadeau", "title": "Carte cadeau"}
        self.assertEqual(Q.page_label(p), "page \u201cCarte cadeau\u201d")


if __name__ == "__main__":
    unittest.main()
