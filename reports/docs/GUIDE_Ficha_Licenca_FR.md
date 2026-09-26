# Guide FR — construire la page « Ficha da Licença » (détail par licence)

Interface Power BI en français. Noms des champs = tels quels (portugais).
On suppose que les tables `Conditions` et `Autocontrolo` sont déjà chargées et
reliées à `Licenses` par `Nº TUA` (sinon, voir la fin du document).

---

## 0. Préparer la page

1. En bas, **+** → nouvelle page → renommer **Ficha da Licença**.
2. Cliquer sur une zone vide de la page. Dans **Visualisations**, section
   **Extraire** (= drill-through), glisser `Licenses[Nº TUA]` dans
   **« Ajouter des champs d'extraction ici »**.
3. Un bouton **flèche retour** apparaît en haut à gauche — le laisser.
4. Clic droit sur l'onglet de la page → **Masquer la page**.

---

## 1. Le bandeau titre (identité de la licence)

1. Clic sur zone vide → insérer une **Carte** (icône `123`).
2. Champ **Champs** → `Licenses[Estabelecimento]`.
3. La placer en haut, l'élargir sur toute la largeur.
4. Mise en forme (pinceau) → **Effet visuel → Étiquette de l'appel** : taille ~20,
   gras. **Étiquette de catégorie** : désactivée.
5. Optionnel : sous le nom, insérer une **Zone de texte** et y glisser (via une
   petite Carte à côté) `Nº TUA`, `TUA/LURH`, `REGIÃO`.

---

## 2. Le bloc des 4 cartes — Estado / Validade / Dias / Tratamento

C'est **4 cartes séparées**, une par valeur (pas une seule carte). Pour chacune :

1. Cliquer sur zone vide → insérer une **Carte** (icône `123`).
2. Glisser **un seul** champ dedans, puis la redimensionner en petit rectangle.
3. Les aligner côte à côte sous le bandeau.

| Carte | Champ à mettre | Table |
|-------|----------------|-------|
| **Estado** | `Estado da Licença` | Licenses |
| **Validade** | `Data de Validade` | Licenses |
| **Dias p/ expirar** | mesure `Dias até Expirar` | Medidas |
| **Tratamento** | `Nível de tratamento` | Licenses |

Astuce titre de chaque carte : Mise en forme → **Étiquette de catégorie**
activée (affiche le nom du champ « Estado da Licença »…), ou désactive-la et
mets un titre propre via **Titre** (« Estado », « Validade », « Dias p/
expirar », « Tratamento »).

> Pourquoi la carte « Dias p/ expirar » utilise une **mesure** et pas une
> colonne : le nombre de jours se recalcule chaque jour par rapport à
> aujourd'hui. La mesure `Dias até Expirar` existe déjà dans la table `Medidas`
> (négatif = déjà caducada).

### Couleur selon l'état (la carte Estado en vert/orange/rouge)

1. Sélectionner la carte **Estado** → Mise en forme → **Étiquette de l'appel →
   Couleur** → **fx** (mise en forme conditionnelle).
2. **Style de format = Valeur du champ** → champ = `Dim_Estado[Cor]`.
   (Cette colonne contient déjà #1E8449 vert, #FF911B orange, #C0392B rouge,
   #95A5A6 gris.) → la valeur se colore automatiquement selon l'état.

---

## 3. Tableau « Paramètres et VLE »

1. Insérer une **Table**.
2. Depuis `Conditions`, glisser dans **Colonnes**, dans l'ordre :
   `Parâmetro`, `VLE`, `VLE mín`, `VLE máx`, `Carga máx. admissível (kg/dia)`,
   `Frequência de amostragem`, `Tipo de amostragem`, `Legislação aplicável`.
3. La placer à gauche, sous les cartes.

---

## 4. Tableau « Autocontrolo » (plan d'échantillonnage)

1. Insérer une **Table**.
2. Depuis `Autocontrolo` : `Local de amostragem`, `Parâmetro`,
   `Frequência de amostragem`, `Tipo de amostragem`, `Nº análises requeridas`.
3. La placer à droite du tableau précédent.

> Normal qu'elle soit vide pour une licence **LURH** : l'autocontrôle n'existe
> que pour les TUA dans le flux combiné.

---

## 5. Matrice « Critères à respecter par paramètre »

1. Insérer une **Matrice**.
2. **Lignes** → `Conditions[Parâmetro]`.
3. **Valeurs** → les 7 colonnes de critères de `Conditions` :
   `média mensal ≤ VLE`, `média anual ≤ VLE`, `≤ 100% VLE (dobro)`,
   `≤ 150% VLE`, `≤ uma ordem de grandeza do VLE`,
   `Gama de valores (intervalo)`, `Quadro III DL 152/97 (borlas)`.
4. Pour chaque valeur : flèche du champ → **Ne pas résumer**.
5. Afficher un ✓ au lieu de 1 : sélectionner la matrice → Mise en forme →
   **Éléments de cellule → Icônes → activer** → **Mise en forme conditionnelle**
   sur chaque colonne → règle : si valeur = 1 → ✓ vert ; sinon vide.

---

## 6. Faire fonctionner le clic depuis la page « Visão Geral »

- Sur **Visão Geral**, clic droit sur une ligne de licence (ou une bulle de la
  carte) → **Extraire → Ficha da Licença**. La page s'ouvre filtrée sur cette
  licence.
- Optionnel — bouton visible : **Insérer → Boutons → Vide**, texte « Ver ficha » ;
  puis **Action → Type = Extraire → Destination = Ficha da Licença**.

---

## 7. Vérifier

- **Accueil → Actualiser**.
- Tester une licence **TUA**, une **LURH**, une **caducada** : les cartes et les
  tableaux doivent changer, la matrice de critères doit montrer les paramètres
  de cette licence.

---

## (Si besoin) Charger + relier Conditions et Autocontrolo

1. **Accueil → Obtenir des données → Excel** → `Power BI Data\Conditions.xlsx`
   → feuille **Conditions** → **Transformer les données**. Régler les types
   numériques (`VLE mín`, `VLE máx` → Nombre décimal). Renommer la requête
   **Conditions**. Idem pour `Autocontrolo.xlsx`. **Fermer et appliquer**.
2. Vue **Modèle** : glisser `Conditions[Nº TUA]` sur `Licenses[Nº TUA]`, puis
   `Autocontrolo[Nº TUA]` sur `Licenses[Nº TUA]`. Cardinalité **Plusieurs à un**,
   sens du filtre **unique**.
