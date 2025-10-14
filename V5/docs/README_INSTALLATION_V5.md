# 🚀 INSTALLATION V5 - Gestion Prêts ESAD

## 📦 FICHIERS LIVRÉS

**3 Modules VBA:**
- `Module1.bas` - Lancement application
- `M_Core.bas` - Fonctions communes
- `M_Business.bas` - Logique métier

**6 UserForms:**
- `MainMenu.frm` / `.frx` - Menu principal
- `LoanHub.frm` / `.frx` - Hub central gestion prêts
- `CreateLoan.frm` / `.frx` - Création prêt
- `ReturnLoan.frm` / `.frx` - Retours (4 méthodes)
- `ManageData.frm` / `.frx` - Gestion articles/emprunteurs
- `Dashboard.frm` / `.frx` - Statistiques

---

## ⚙️ INSTALLATION

### 1. Préparer fichier Excel

**Option A: Partir de v2 existante**
```
1. Ouvrir pret_mat_v2.xlsm
2. Fichier > Enregistrer sous > pret_mat_v5.xlsm
3. Ouvrir VBA (Alt+F11)
```

**Option B: Nouveau fichier**
```
1. Créer nouveau classeur Excel
2. Enregistrer en .xlsm
3. Créer feuilles: accueil, emprunteurs, prets, articles, service, fonction, tech, résultat
4. Créer tables nommées: Tableau1, Tableau10, Tableau4
```

### 2. Importer Modules

```
VBA Editor > Clic droit sur VBAProject > Importer
Sélectionner les 3 fichiers .bas:
- Module1.bas
- M_Core.bas
- M_Business.bas
```

### 3. Importer UserForms

```
VBA Editor > Clic droit sur VBAProject > Importer
Sélectionner les 6 fichiers .frm (les .frx seront importés automatiquement):
- MainMenu.frm
- LoanHub.frm
- CreateLoan.frm
- ReturnLoan.frm
- ManageData.frm
- Dashboard.frm
```

### 4. Vérifier compilation

```
VBA Editor > Menu Debug > Compiler VBAProject

Si erreurs:
- Vérifier que toutes les feuilles existent
- Vérifier que Tableau1, Tableau10, Tableau4 existent
```

### 5. Tester

```
1. Fermer VBA Editor
2. Fermer Excel
3. Rouvrir pret_mat_v5.xlsm
4. Le menu MainMenu devrait s'afficher automatiquement
```

---

## ✅ VÉRIFICATIONS POST-INSTALLATION

- [ ] Menu principal s'affiche au démarrage
- [ ] Boutons du menu fonctionnent
- [ ] Hub prêts accessible
- [ ] Export inventaire fonctionne
- [ ] Dashboard affiche statistiques

---

## 🆕 NOUVELLES FONCTIONNALITÉS V5

**1. Export Inventaire Complet**
- Bouton dans menu principal
- Feuille Excel avec tous les articles
- Coloration automatique (vert=dispo, orange=prêté)

**2. Retours Groupés (3 méthodes)**
- **Tout retourner:** 1 clic pour tout retourner
- **Cochage:** Sélection multiple avec checkboxes
- **Scan chaîne:** Scan QR à la chaîne

**3. Dashboard Statistiques**
- Stats globales
- Alertes (prêts >15j, >30j)
- Top 10 articles les plus prêtés

**4. Interface Moderne**
- Design épuré
- Couleurs cohérentes
- Navigation intuitive

---

## 🐛 DÉPANNAGE

**Le menu ne s'affiche pas au démarrage:**
- Vérifier que Module1.bas est importé
- Vérifier que Auto_Open() existe
- Macro désactivée? Activer les macros

**Erreur "Table introuvable":**
- Créer les tables nommées manquantes
- Vérifier orthographe exacte

**Erreur compilation:**
- Vérifier que tous les modules sont importés
- Vérifier que tous les UserForms sont importés

**Contrôles non visibles:**
- C'est normal! Tous les contrôles sont créés par code
- Vérifier UserForm_Initialize() dans chaque .frm

---

## 📞 SUPPORT

Florian Limmelette
florian.limmelette@esad-orleans.fr
06 62 09 72 19

---

**Version:** 5.0  
**Date:** 14 octobre 2025
