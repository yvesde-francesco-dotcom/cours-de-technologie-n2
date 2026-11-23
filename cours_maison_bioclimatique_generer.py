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


def titre(txt, size=16, color=VERT, space_before=16, space_after=6, center=False):
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


def para(txt, italic=False):
    p = doc.add_paragraph()
    r = p.add_run(txt)
    r.italic = italic
    p.paragraph_format.space_after = Pt(6)
    return p


def puce(txt, gras_avant=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    if gras_avant:
        r = p.add_run(gras_avant)
        r.bold = True
    p.add_run(txt)
    return p


def encadre(txt, prefix="À retenir : "):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(prefix)
    r.bold = True
    r.font.color.rgb = VERT
    p.add_run(txt)
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
r = s.add_run("Cours de Technologie — Classe de 5e")
r.font.size = Pt(14)
r.font.color.rgb = GRIS
r2 = s.add_run("\nThème : vivre dans un habitat confortable et respectueux de l'environnement")
r2.font.size = Pt(11)
r2.font.color.rgb = GRIS

obj = doc.add_paragraph()
obj.paragraph_format.space_before = Pt(50)
obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = obj.add_run("Objectifs du cours")
r.bold = True
r.font.color.rgb = VERT
r.font.size = Pt(12)
for objectif in [
    "Comprendre ce qu'est une maison bioclimatique",
    "Connaître les règles d'orientation d'une maison",
    "Comprendre ce qu'est l'énergie grise",
    "Savoir comment l'isolation garde la chaleur (ou la bloque)",
    "Être capable de concevoir une maison économe en énergie",
]:
    p = doc.add_paragraph(style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(objectif)

doc.add_page_break()

# ================= 1. Introduction =================
titre("1. Qu'est-ce qu'une maison bioclimatique ?", size=17, space_before=0)
para(
    "« Bio » signifie la vie et la nature ; « climatique » rappelle le climat et le "
    "soleil. Une maison bioclimatique est donc une maison pensée et construite pour "
    "limiter sa consommation d'énergie : elle utilise au maximum la chaleur et la "
    "lumière gratuites du soleil, au lieu de tout demander au chauffage et à l'éclairage "
    "électrique."
)
para(
    "La maison s'adapte ainsi à son environnement et non l'inverse. Sa conception repose "
    "sur l'étude de plusieurs éléments : le choix du terrain, l'orientation, la forme, "
    "l'isolation, la ventilation et le choix du système de chauffage."
)
encadre(
    "Le bioclimatisme consiste à exploiter gratuitement l'environnement (soleil, vent, "
    "végétation) pour optimiser le confort et limiter les besoins en chauffage, en "
    "refroidissement et en éclairage."
)

image("page-0001-image-0001.jpg", 12,
      "Schéma d'une maison bioclimatique : façade sud largement vitrée, casquette "
      "protectrice, arbres caduc au sud et persistant au nord, capteurs solaires, "
      "ventilation et déperditions limitées.")

# ================= 2. Orientation =================
titre("2. L'orientation : la règle d'or du bioclimatisme")
para(
    "L'orientation de la maison et la répartition des pièces selon les points cardinaux "
    "sont essentielles, car les apports solaires sont gratuits mais inégaux selon la "
    " saison et la façade."
)
titre("Répartition des pièces", size=13, color=BLEU, space_before=8, space_after=4)
puce("Au SUD : les pièces de vie (séjour, salon) avec de grandes baies vitrées, pour être "
     "chauffées par le soleil ; une casquette (débord de toit) protège des rayons "
     "estivaux trop hauts tout en laissant passer les rayons bas de l'hiver.",
     "Sud : ")
puce("au NORD : les pièces ne nécessitant pas de chaleur (atelier, garage, buanderie), "
     "qui forment une zone tampon contre le froid.",
     "Nord : ")
titre("La végétation au service de l'orientation", size=13, color=BLEU, space_before=8, space_after=4)
puce("Un arbre à feuillage caduc planté au sud laisse passer le soleil en hiver et le "
     "masque de son feuillage en été.")
puce("Un arbre persistant planté au nord protège la maison du vent froid.")
image("page-0003-image-0001.jpg", 5, "Rose des vents : repère d'orientation des façades.")

# ================= 3. Forme =================
titre("3. La forme de la maison : la compacité")
para(
    "Imagine : quand tu es dehors par grand froid, il est plus facile de garder ta "
    "chaleur en rentrant les bras le long du corps (tu présentes une petite surface) que "
    "bras et jambes écartés. Une maison, c'est pareil !"
)
para(
    "À surface habitable égale, la quantité de murs en contact avec l'air extérieur "
    "détermine les déperditions de chaleur. Plus la maison est compacte, moins elle "
    "se refroidit."
)
encadre(
    "Deux maisons de 144 m² : une maison de 12 m × 12 m (carrée) présente 48 m de "
    "périmètre de murs, contre 49,2 m pour une maison de 15 m × 9,6 m (rectangulaire). "
    "La maison carrée, plus compacte, sera donc plus facile à chauffer."
)

# ================= 4. Matériaux =================
doc.add_page_break()
titre("4. Le choix des matériaux", size=17, space_before=0)
para(
    "Un matériau de construction ne coûte pas seulement en euros : il a aussi un coût "
    "en énergie, appelé énergie grise."
)
titre("L'énergie grise", size=13, color=BLEU, space_before=8, space_after=4)
encadre(
    "L'énergie grise est l'énergie nécessaire à l'extraction, à la fabrication et au "
    "transport d'un matériau.", prefix="Définition : "
)
titre("Exemple de classement", size=13, color=BLEU, space_before=8, space_after=4)
image("page-0004-image-0001.jpg", 11,
      "Énergie grise comparée de trois matériaux de construction.")
puce("La paille : énergie grise la plus faible (matériau local, peu transformé).", "1. ")
puce("Le bois : énergie grise faible (renouvelable, peu de transformation).", "2. ")
puce("Le parpaing : énergie grise la plus élevée (fabrication industrielle, "
     "cimentiers énergivores).", "3. ")
para(
    "Une maison écologique privilégie donc les matériaux isolants, non polluants et à "
    "faible énergie grise.", italic=False
)

# ================= 5. Isolation =================
titre("5. L'isolation thermique")
titre("Où la chaleur s'échappe-t-elle ?", size=13, color=BLEU, space_before=8, space_after=4)
para(
    "La chaleur contenue à l'intérieur d'une maison s'échappe par le toit (première source "
    "de déperdition), puis par les murs et les vitres."
)
titre("Le rôle de l'isolant", size=13, color=BLEU, space_before=8, space_after=4)
para(
    "Un isolant thermique sert à empêcher la chaleur de s'échapper de la maison en hiver "
    "ou, au contraire, à éviter qu'elle n'entre à l'intérieur en été."
)
encadre(
    "Pour être le plus efficace, l'isolant doit être placé à l'extérieur de la structure "
    "(isolation par l'extérieur), afin d'éviter les ponts thermiques : ces points de "
    "contact qui laissent la chaleur traverser le mur."
)
image("page-0007-image-0001.jpg", 10,
      "En hiver : la chaleur intérieure se déplace vers l'extérieur froid à travers "
      "la paroi.")
image("page-0007-image-0002.jpg", 10,
      "Avec une isolation placée à l'extérieur du mur, le flux de chaleur est bloqué : "
      "le froid n'atteint plus le mur intérieur et il n'y a pas de pont thermique.")

# ================= 6. Saisons et apports solaires =================
doc.add_page_break()
titre("6. Un confort toute l'année : hiver / été", size=17, space_before=0)
para(
    "La conception bioclimatique joue le rôle d'un régulateur de température naturel, "
    "sans consommer d'énergie :"
)
puce("en hiver, le soleil bas traverse les baies vitrées du sud et chauffe les pièces de "
     "vie ; l'isolation conserve cette chaleur à l'intérieur.",
     "Hiver : ")
puce("en été, le soleil haut ne passe pas : la casquette et le feuillage des arbres caducs "
     "ombragent les vitres, et l'isolation empêche la chaleur extérieure de pénétrer.",
     "Été : ")
image("page-0007-image-0003.jpg", 11,
      "En été : casquette et végétation bloquent les rayons solaires hauts ; la maison "
      "reste fraîche.")
image("page-0007-image-0004.jpg", 11,
      "Comportement saisonnier de la maison bioclimatique selon l'inclinaison du soleil.")

# ================= 7. Ventilation et énergie =================
titre("7. Ventilation et équipements")
puce("la ventilation renouvelle l'air intérieur tout en limitant la consommation de "
     "chauffage (récupération de chaleur de l'air extrait).",
     "Ventilation : ")
puce("panneaux photovoltaïques pour l'électricité, panneaux solaires thermiques pour "
     "l'eau chaude.",
     "Solaire : ")
puce("chauffage au bois (énergie renouvelable) pour compléter les apports gratuits.",
     "Chauffage : ")
puce("récupération de l'eau de pluie pour l'arrosage et les sanitaires.",
     "Eau : ")

# ================= 8. Quiz =================
titre("8. Teste tes connaissances")
para("Réponds d'abord, puis vérifie avec le cours :")
quiz = [
    "Quelle est la façade à vitrer en priorité, et pourquoi ?",
    "Où place-t-on le garage, la buanderie et l'atelier ?",
    "Quel arbre plante-t-on au sud, et quelle protection apporte-t-il selon la saison ?",
    "Classe du plus petit au plus grand : énergie grise du bois, du parpaing, de la paille.",
    "Cite les trois endroits par lesquels la chaleur s'échappe d'une maison.",
    "Pourquoi l'isolant doit-il être posé à l'extérieur du mur ?",
    "En été, qu'est-ce qui empêche le soleil de surchauffer le séjour ?",
]
for i, q in enumerate(quiz, 1):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(f"{i}. ")
    r.bold = True
    r.font.color.rgb = BLEU
    p.add_run(q)
encadre(
    "Ça y est, tu as tout appris ? Avec ces principes, tu peux maintenant expliquer "
    "pourquoi une maison bioclimatique chauffe moins et vit mieux — et c'est le thème "
    "du programme de 5e : vivre dans un habitat confortable et respectueux de "
    "l'environnement.",
    prefix="Pour aller plus loin : ",
)

# ================= 9. Avantages / limites =================
titre("9. Avantages et limites")
titre("Les avantages", size=13, color=BLEU, space_before=8, space_after=4)
puce("Impact environnemental réduit grâce à une consommation d'énergie moindre.")
puce("Excellent confort thermique toute l'année, sans surconsommation.")
puce("Maisons lumineuses : grandes ouvertures bien placées nécessitant peu d'éclairage "
     "artificiel.")
titre("Les limites", size=13, color=BLEU, space_before=8, space_after=4)
puce("Coûts de construction potentiellement plus élevés que pour une maison classique ; "
     "l'objectif de rentabilité se joue à long terme.")
puce("La maison doit s'adapter à son environnement (terrain, orientation) et non "
     "l'inverse, ce qui contraint parfois le plan et l'implantation.")

# ================= 10. Synthèse =================
titre("10. Synthèse", space_after=4)
encadre(
    "Une maison bioclimatique s'oriente plein sud pour ses pièces de vie, se protège au "
    "nord (pièces tampons, arbres persistants), adopte une forme compacte, utilise des "
    "matériaux à faible énergie grise, s'isole par l'extérieur pour supprimer les ponts "
    "thermiques, maîtrise le soleil grâce aux casquettes et à la végétation, complète avec "
    "des énergies renouvelables et récupère l'eau de pluie."
)
liste = doc.add_paragraph()
r = liste.add_run("Mots-clés : ")
r.bold = True
liste.add_run("bioclimatisme — énergie grise — compacité — pont thermique — isolation "
              "par l'extérieur — casquette — énergie renouvelable — apports solaires "
              "gratuits — déperdition — arbre caduc/persistant")

doc.save("cours_maison_bioclimatique.docx")
print("OK")
