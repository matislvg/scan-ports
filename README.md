# 🔍 Scanner de Ports Réseaux

> Scanner de ports réseau multithreadé avec rapport Excel de qualité audit — écrit en Python.

Cet outil automatise la **découverte d'hôtes** et le **balayage des ports sensibles** sur un réseau, puis génère un rapport `.xlsx` entièrement mis en page et colorisé selon le niveau de criticité de chaque service exposé.

---

## Contexte

Ce script a été développé dans le cadre de mon **stage de cybersécurité** au sein de l'[Entente Valabre](https://www.entente-valabre.com) (Pôle Innovations et Nouvelles Technologies), structure interrégionale de sécurité civile basée en région PACA.

L'objectif était de disposer d'un outil d'audit interne rapide, sans dépendance à des solutions tierces comme Nmap, pour cartographier les ports sensibles exposés sur le réseau de l'organisation et alimenter la démarche de supervision sécurité en place.

---

## Aperçu du rapport généré

| Adresse IP | Statut | Hostname | Port 22 SSH (ÉLEVÉ) | Port 23 Telnet (CRITIQUE) | Port 443 HTTPS (FAIBLE) |
|---|---|---|---|---|---|
| 192.168.1.1 | 🟢 UP | gateway.local | 🟠 OUVERT | 🔴 OUVERT | 🔵 OUVERT |
| 192.168.1.5 | 🟢 UP | workstation-05 | 🟢 FERMÉ | 🟢 FERMÉ | 🟢 FERMÉ |
| 192.168.1.9 | ⚫ DOWN | — | — | — | — |

---

## Fonctionnalités

- **Multithreading** — scan parallèle des hôtes via `ThreadPoolExecutor`, gain de ×10 à ×20 sur les performances
- **Découverte d'hôtes en deux temps** — Ping ICMP, puis fallback TCP 80/443 si l'ICMP est filtré par un pare-feu
- **Gestion CIDR complète** — IP seule, liste d'IPs, sous-réseaux (`192.168.1.0/24`), y compris les `/32`
- **Rapport Excel automatisé** :
  - Toutes les machines du scope présentes, y compris celles en `DOWN`
  - Colorisation logique selon la criticité du service (voir tableau ci-dessous)
  - En-têtes figés, colonnes ajustées, grille masquée
- **CLI complète** — timeouts, workers, fichier de sortie, mode verbose, tout est paramétrable
- **Interruption propre** — `Ctrl+C` exporte les résultats déjà collectés au lieu de tout perdre

---

## Logique de colorisation

| Criticité | Ports concernés | Couleur si OUVERT |
|---|---|---|
| `CRITIQUE` | 21 (FTP), 23 (Telnet), 139/445 (SMB), 3389 (RDP) | 🔴 Rouge |
| `ÉLEVÉ` | 22 (SSH), 25 (SMTP), 5985/5986 (WinRM) | 🟠 Orange |
| `MODÉRÉ` | 53 (DNS), 80 (HTTP) | 🟡 Jaune |
| `FAIBLE` | 443 (HTTPS) | 🔵 Bleu |
| — | Tout port | 🟢 Vert si FERMÉ |
| — | Machine DOWN | ⚫ Gris |

> Un port 443 (HTTPS) ouvert est un comportement normal — il ne sera pas affiché en rouge alarmiste comme dans la plupart des scanners basiques.

---

## Prérequis

- Python **3.9+**
- Droits **administrateur** recommandés (requis pour le Ping ICMP sous Windows)

---

## Installation

```bash
git clone https://github.com/votre-username/scan-ports.git
cd scan-ports
pip install -r requirements.txt
```

---

## Utilisation

```
python scan_ports.py -t <CIBLES> [options]
```

### Options

| Option | Alias | Description | Défaut |
|---|---|---|---|
| `--targets` | `-t` | Cibles à scanner (IP ou CIDR, séparées par des espaces) — **obligatoire** | — |
| `--output` | `-o` | Nom du fichier Excel généré | `rapport_scan.xlsx` |
| `--workers` | `-w` | Nombre de threads parallèles (max recommandé : 50) | `20` |
| `--timeout` | — | Timeout par connexion en secondes | `1.0` |
| `--verbose` | `-v` | Active les logs DEBUG | désactivé |

### Exemples

```bash
# Scanner un sous-réseau complet
python scan_ports.py -t 192.168.1.0/24

# Scanner plusieurs cibles avec un nom de fichier personnalisé
python scan_ports.py -t 10.0.0.1 10.0.0.5 172.16.0.0/28 -o audit_infra.xlsx

# Scan rapide avec plus de threads et timeout réduit
python scan_ports.py -t 192.168.0.0/24 --workers 40 --timeout 0.5

# Mode verbeux pour le débogage
python scan_ports.py -t 10.0.0.0/28 -v
```

---

## Structure du projet

```
network-port-scanner/
├── scan_ports.py   # Script principal
├── requirements.txt    # Dépendances Python
├── .gitignore          # Exclusion des rapports générés
└── README.md
```

---

## Stack technique

| Module | Rôle |
|---|---|
| `socket` | Connexions TCP pour le scan de ports |
| `subprocess` + `platform` | Ping ICMP cross-platform (Windows / Linux / macOS) |
| `ipaddress` | Résolution et décomposition des plages CIDR |
| `concurrent.futures` | Parallélisation via ThreadPoolExecutor |
| `pandas` | Construction et export du DataFrame |
| `openpyxl` | Mise en forme avancée du rapport Excel |
| `argparse` | Interface CLI |
| `logging` | Logs horodatés avec niveau configurable |

---

## ⚠️ Avertissement légal

Cet outil est développé à des fins **éducatives et d'audit de sécurité interne**.  
Le scan de ports sur un réseau sans l'autorisation explicite de son propriétaire est **strictement interdit** et peut être constitutif d'une infraction pénale selon la législation applicable (notamment l'article 323-1 du Code pénal en France).  
L'auteur décline toute responsabilité quant à une utilisation non autorisée ou malveillante de ce script.

---

## Licence

Ce projet a été réalisé dans le cadre d'un stage à l'Entente Valabre.  
Réutilisation soumise à autorisation.
