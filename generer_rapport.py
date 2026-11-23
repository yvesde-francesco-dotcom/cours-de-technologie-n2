# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

IMG = "images"

VERT = RGBColor(0x2E, 0x7D, 0x32)
BLEU = RGBColor(0x1A, 0x55, 0x8A)
GRIS = RGBColor(0x55, 0x55, 0x55)

doc = Document()

normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(11)


def titre(txt, size=16, color=VERT, space_before=14, space_after=8, center=False):
    p = doc.add_paragraph()
    r = p.add_run(txt)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p


def question(txt):
    p = doc.add_paragraph()
    r = p.add_run(txt)
    r.bold = True
    r.font.size = Pt(11.5)
    r.font.color.rgb = BLEU
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    return p


def reponse(txt):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("Réponse : ")
    r.bold = True
    r.font.color.rgb = VERT
    p.add_run(txt)
    return p


def puce(txt):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(txt)
    p.paragraph_format.space_after = Pt(2)
    return p


def image(fichier, largeur_cm, legende):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(f"{IMG}/{fichier}", width=Cm(largeur_cm))
    lp = doc.add_paragraph()
    lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = lp.add_run(legende)
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = GRIS
    lp.paragraph_format.space_after = Pt(10)


# ================= Page de titre =================
titre("La maison bioclimatique", size=28, space_before=140, space_after=2, center=True)
s = doc.add_paragraph()
s.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = s.add_run("Cours — Concevoir une habitation économe en énergie")
r.font.size = Pt(14)
r.font.color.rgb = GRIS

intro = doc.add_paragraph()
intro.alignment = WD_ALIGN_PARAGRAPH.CENTER
intro.paragraph_format.space_before = Pt(180)
r.font.color.rgb = GRIS
intro.text = ""
# ================= 1. Définition =================
titre("1. Définition", size=17, space_before=6)
p = doc.add_paragraph(
    "Une maison « bioclimatique » désigne une habitation pensée et conçue de manière "
    "à limiter la consommation d'énergie. L'idée est de profiter au maximum des apports "
    "naturels de chaleur et de lumière du soleil. Ce principe repose sur l'étude de "
    "différents éléments relatifs à la conception de l'habitation :"
)
p.paragraph_format.space_after = Pt(6)
puce("Le choix du terrain ;")
puce("L'orientation de la maison (pour profiter au maximum des apports solaires) ;")
puce("L'isolation (pour réduire au maximum les besoins de chauffage) ;")
puce("La ventilation (pensée, là encore, pour limiter la consommation de chauffage) ;")
puce("Le choix d'un système de chauffage respectueux de l'environnement et économique.")

image("page-0001-image-0001.jpg", 12,
      "Schéma d'une maison bioclimatique : façade sud exposée au soleil, casquette "
      "protectrice en été, arbres caduc au sud et persistant au nord, capteurs solaires, "
      "ventilation et déperditions limitées.")

# ================= 2. Avantages / Inconvénients =================
titre("2. Avantages et inconvénients")
sous = titre("Avantages", size=13, color=VERT, space_before=6, space_after=4)
puce("Impact environnemental réduit grâce à une consommation d'énergie moindre "
     "(chauffage, refroidissement et éclairage optimisés).")
puce("Excellent confort thermique toute l'année.")
puce("Maisons lumineuses : grandes ouvertures bien placées et orientées, nécessitant "
     "peu d'éclairage artificiel.")

titre("Inconvénients", size=13, color=VERT, space_before=8, space_after=4)
puce("Coûts de construction potentiellement plus élevés que pour une maison classique ; "
     "l'objectif reste de réaliser des économies à long terme.")
puce("La maison bioclimatique s'adapte à son environnement, et non l'inverse, ce qui "
     "contraint le choix du terrain et du plan.")

# ================= 3. Conception d'une maison écologique =================
doc.add_page_break()
titre("3. Conception d'une maison écologique", size=17, space_before=0)
question("Quelles sont les caractéristiques d'une maison écologique ?")
reponse("Utilisation de matériaux isolants non polluants, en évitant les matériaux "
        "qui consomment trop d'énergie grise.")
reponse("Utilisation de panneaux photovoltaïques pour l'électricité et de panneaux "
        "solaires pour l'eau chaude.")
reponse("Utilisation de baies vitrées pour le chauffage.")
reponse("Utilisation d'énergies renouvelables pour le chauffage (ex. le bois).")
reponse("Récupération de l'eau de pluie.")

# ================= 4. Aménagement =================
titre("4. Aménagement d'une maison écologique")
question("Indique face aux points cardinaux les pièces qu'il convient d'orienter.")
image("page-0003-image-0001.jpg", 5.5, "Rose des vents (points cardinaux).")
reponse("Au NORD : atelier, garage, buanderie — les pièces ne nécessitant pas de chaleur.")
reponse("Au SUD : séjour et salon, avec une baie vitrée pour être chauffés par le soleil, "
        "protégée par une casquette.")

question("Deux maisons de surface habitable identique (144 m²) : laquelle pourra, a priori, "
         "être chauffée le plus facilement ? Pourquoi ?")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(f"{IMG}/page-0003-image-0002.jpg", width=Cm(6))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Maison de 15 m × 9,6 m").italic = True
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(f"{IMG}/page-0003-image-0003.jpg", width=Cm(6))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Maison de 12 m × 12 m")
r.italic = True
p.paragraph_format.space_after = Pt(8)
reponse("La maison carrée pourra être chauffée plus facilement, car elle présente moins de "
        "surface extérieure en contact avec l'air que la maison rectangulaire "
        "(48 m de mur contre 49,2 m).")

# ================= 5. Choix des matériaux =================
doc.add_page_break()
titre("5. Choix des matériaux de construction", size=17, space_before=0)
question("Qu'est-ce que l'énergie grise ?")
reponse("L'énergie grise est l'énergie nécessaire à l'extraction, à la fabrication et au "
        "transport du matériau utilisé.")

question("Classe ces trois matériaux dans l'ordre croissant de leur énergie grise : "
         "le bois, le parpaing et la paille.")
image("page-0004-image-0001.jpg", 11,
      "Énergie grise comparée : la paille (la plus faible), le bois, puis le parpaing "
      "(la plus élevée).")
reponse("Ordre croissant : la paille, puis le bois, puis le parpaing.")

# ================= 6. Isolation thermique =================
titre("6. L'isolation thermique d'une maison")
question("Par où s'échappe la chaleur contenue à l'intérieur d'une maison ?")
reponse("Elle s'échappe par le toit, les murs et les vitres.")

question("À quoi sert un isolant thermique ?")
reponse("Un isolant thermique sert à empêcher la chaleur de s'échapper de la maison ou, "
        "au contraire, à éviter qu'elle n'entre à l'intérieur.")

question("Où doit-il être placé pour être le plus efficace ?")
reponse("Il doit être placé à l'extérieur de la maison pour éviter les ponts thermiques.")

question("Explique pourquoi, en figurant par des flèches le chemin de la chaleur sur "
         "chacun des deux dessins suivants.")
image("page-0007-image-0001.jpg", 11,
      "En hiver : la chaleur intérieure (chaud) traverse le mur non isolé vers l'extérieur "
      "(froid).")
image("page-0007-image-0002.jpg", 11,
      "Avec un isolant placé à l'extérieur du mur : le flux de chaleur est bloqué, il n'y a "
      "plus de pont thermique.")
image("page-0007-image-0003.jpg", 11,
      "En été : le soleil chauffe la maison, la casquette et l'isolation limitent la surchauffe.")
image("page-0007-image-0004.jpg", 11,
      "Comportement de la maison bioclimatique en hiver et en été.")

doc.save("la_maison_bioclimatique_presentation.refaite.docx")
print("OK")
