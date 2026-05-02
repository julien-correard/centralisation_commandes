# 📦 Centralisation de commandes clients

Script Python permettant de **consolider automatiquement des commandes clients** à partir de fichiers Excel hétérogènes.

---

## 🎯 Objectif

Agréger des données issues de plusieurs fichiers `.xlsx` (clients) dans un fichier unique (`Commandes.xlsx`), malgré des formats non uniformes.

---

## 📌 Contexte

Projet réalisé dans le cadre d’une reconversion vers le développement.

Le script est actuellement utilisé en production dans une entreprise pour automatiser la centralisation de commandes clients, réduisant significativement le traitement manuel et les erreurs associées.

---

## 🧠 Choix techniques

### 🔌 Configuration externalisée

* Correspondance des clients définie dans :

  ```
  script/Correspondance_clients.csv
  ```
* Association entre :

  * nom complet du client
  * abréviation utilisée dans le fichier de sortie

➡️ Ajout ou modification de clients **sans modification du code**

---

### 🔍 Parcours dynamique des données

* Aucun emplacement fixe imposé dans les fichiers
* Lecture complète des tableaux
* Identification des produits par leur nom

➡️ Permet de traiter des fichiers **non standardisés**

---

### ⚙️ Logique générique

* Ajout de nouveaux produits sans modification du script
* Adaptation à différents formats clients
* Séparation claire entre données et logique

---

### 🛡️ Robustesse

Le script adopte une approche **fail-fast** : en cas d’anomalie (données incohérentes, correspondance manquante, fichier invalide), le traitement est interrompu afin d’éviter toute erreur silencieuse ou corruption des données.

➡️ Choix assumé : privilégier la fiabilité et la traçabilité plutôt que la continuité à tout prix.

---

## 🛠️ Technologies

* Python 3
* Manipulation de fichiers Excel (`.xlsx`)
* Fichier CSV pour la configuration

---

## 🎓 Ce que démontre le projet

* Traitement de données réelles et imparfaites
* Conception simple mais évolutive
* Séparation entre logique métier et configuration
* Approche orientée robustesse plutôt que rigidité

--

## A améliorer

La gestion des erreurs est actuellement fonctionnelle mais pourrait être rendue plus idiomatique en Python.
Notamment, la séparation entre validation, logique métier et interaction utilisateur pourrait être améliorée (par exemple via des exceptions personnalisées plutôt que des appels directs à input() ou sys.exit()).

Ce projet étant mon premier en Python, il a été construit de manière progressive, avec un apprentissage en cours de développement. Une refactorisation future permettrait d'améliorer la structure et la maintenabilité sur ces aspects.
