# Tableau de bord Power BI – Suivi hebdomadaire des mises en place de crédits

**Contexte :** Direction du Crédit – NSIA Banque Sénégal

## Problème : 
Chaque semaine, la Direction du Crédit de NSIA Banque Sénégal reçoit un fichier Excel listant
les crédits et engagements mis en place. Elle souhaite en tirer une vue claire de la production :
combien a été engagé, sur quels produits, pour quels clients, dans quelles agences et sur quelles
durées.Ces tableaux de bord étaient auparavant construits manuellement par les équipes, ce qui
demandait **4 à 5 heures** par semaine.

## Solution : 
J’ai conçu un rapport Power BI automatisé qui produit le même résultat en **moins de 5 minutes**.

## Objectifs du projet :
- automatiser la production du tableau de bord hebdomadaire
- appliquer des règles de classification identiques d’une semaine à l’autre
- rendre visibles les anomalies de données plutôt que de les masquer.

### Fonctionnalités principales
- Consolidation automatique de plusieurs fichiers Excel (DIRECTS + HORS BILAN)
- Classification des produits de crédit selon des règles métier (CMTA, ESC1, CCTN, GAR, GRT…)
- Calcul des tranches de maturité
- Enrichissement avec tables de référence (Agences + Segment client)
- Contrôles qualité explicites (« À CLASSER », « NON CLASSÉ », « INCONNU »)
- Indicateurs de concentration (part des 3 premiers clients)
- Vue photographie de la période + vue évolution hebdomadaire

### Stack technique
- **Power Query** : extraction, nettoyage, classification et consolidation
- **DAX** : mesures (Engagements Directs, Hors-bilan, Top 3 Clients, etc.)
- **Power BI** : restitution (2 pages principales)

  ### Impact
| Avant | Après |
|-------|-------|
| 4 à 5 heures / semaine | moins de 5 minutes |
| Règles variables selon la personne | Règles formalisées et identiques |
| Anomalies souvent masquées | Anomalies visibles et tracées |

## Architecture
1. Sources : fichiers hebdomadaires + tables de référence
2. Power Query : lecture, classification, contrôles qualité
3. Modèle + DAX
4. Rapport (2 pages)

## Confidentialité
Les données clients et les fichiers sources réels ne peuvent pas être partagés pour des raisons de confidentialité.  
Les captures d’écran présentées sont anonymisées.
