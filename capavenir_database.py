"""
=============================================================================
CapAvenir CMR — Module Base de Données (SQLite)
=============================================================================
Fichier  : capavenir_database.py
Version  : 1.0
Auteur   : CapAvenir CMR — Mémoire ENS Filière Informatique Niveau 5

Description :
    Ce module gère toute la persistance des données de l'application
    d'orientation scolaire CapAvenir CMR.

    3 tables principales :
      • orientations_confirmees  — Dossiers validés définitivement
      • orientations_en_attente  — Dossiers en attente (T2 ou T3 requis)
      • orientations_probation   — Dossiers en probation (suivi requis)

    + 1 table de logs :
      • historique_statuts        — Historique des changements de statut

Utilisation dans streamlit_run_app.py :
    from capavenir_database import CapAvenirDB

    db = CapAvenirDB()                     # initialise / ouvre la BDD
    db.sauvegarder_dossier(session_state)  # sauvegarde automatique selon statut
    dossiers = db.lister_dossiers("confirme")
    dossier  = db.rechercher_eleve("MBALLA", "Jean")
    db.supprimer_dossier(dossier_id)
    db.close()
=============================================================================
"""

import sqlite3
import json
import os
from datetime import datetime
from typing import Optional


# =============================================================================
# CONSTANTES
# =============================================================================
DB_FILE = "capavenir_cmr.db"   # Chemin du fichier SQLite (même dossier que l'app)

STATUTS_VALIDES = ("confirme", "attente", "probation", "revise", "indetermine")

# Mapping statut → table de destination
TABLE_PAR_STATUT = {
    "confirme":    "orientations_confirmees",
    "revise":      "orientations_confirmees",   # Révisé = décision finale aussi
    "attente":     "orientations_en_attente",
    "probation":   "orientations_probation",
    "indetermine": "orientations_en_attente",   # Indéterminé → file d'attente
}


# =============================================================================
# CLASSE PRINCIPALE
# =============================================================================
class CapAvenirDB:
    """
    Interface complète avec la base de données SQLite de CapAvenir CMR.

    Exemple d'utilisation :
        db = CapAvenirDB()
        db.sauvegarder_dossier(st.session_state)
        resultats = db.lister_dossiers("confirme")
        db.close()
    """

    def __init__(self, db_path: str = DB_FILE):
        """
        Ouvre (ou crée) la base de données et initialise les tables.

        Args:
            db_path: Chemin vers le fichier .db (défaut : capavenir_cmr.db)
        """
        self.db_path = db_path
        self.conn    = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row   # Résultats accessibles par nom de colonne
        self._activer_foreign_keys()
        self._creer_tables()

    # -------------------------------------------------------------------------
    # INITIALISATION
    # -------------------------------------------------------------------------
    def _activer_foreign_keys(self):
        """Active le support des clés étrangères dans SQLite."""
        self.conn.execute("PRAGMA foreign_keys = ON")

    def _creer_tables(self):
        """Crée toutes les tables si elles n'existent pas encore."""
        cursor = self.conn.cursor()

        # ----- Table des orientations CONFIRMÉES (+ RÉVISÉES) -----
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS orientations_confirmees (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            -- Identité de l'élève
            nom                 TEXT    NOT NULL,
            prenom              TEXT    NOT NULL,
            age                 INTEGER,
            sexe                TEXT,
            lycee               TEXT,
            choix_personnel     TEXT,
            projet_pro          TEXT,
            revenu_famille      TEXT,

            -- Résultat de l'orientation
            serie_finale        TEXT    NOT NULL,          -- 'C' ou 'A'
            statut              TEXT    NOT NULL,          -- 'confirme' ou 'revise'
            score_confiance     INTEGER,                   -- 0-100 %
            trimestre_decision  TEXT,                      -- 'T1', 'T2' ou 'T3'

            -- Scores tests psychotechniques (bruts)
            d48_brut            REAL,
            krx_brut            REAL,
            meca_brut           REAL,
            bv11_brut           REAL,
            prc_brut            REAL,

            -- Scores étalonnés
            d48_etal            REAL,
            krx_etal            REAL,
            meca_etal           REAL,
            bv11_etal           REAL,
            prc_etal            REAL,

            -- Aptitudes calculées
            SA_brut             REAL,
            SA_etal             REAL,
            LA_brut             REAL,
            LA_etal             REAL,

            -- Notes scolaires T1
            maths_t1            REAL,
            sci_phy_t1          REAL,
            svt_t1              REAL,
            francais_t1         REAL,
            histgeo_t1          REAL,
            anglais_t1          REAL,
            moy_sci_t1          REAL,
            moy_lit_t1          REAL,

            -- Notes scolaires T2 (si disponibles)
            maths_t2            REAL,
            sci_phy_t2          REAL,
            svt_t2              REAL,
            francais_t2         REAL,
            histgeo_t2          REAL,
            anglais_t2          REAL,
            moy_sci_t2          REAL,
            moy_lit_t2          REAL,

            -- Notes scolaires T3 (si disponibles)
            maths_t3            REAL,
            sci_phy_t3          REAL,
            svt_t3              REAL,
            francais_t3         REAL,
            histgeo_t3          REAL,
            anglais_t3          REAL,
            moy_sci_t3          REAL,
            moy_lit_t3          REAL,

            -- Enrichissement
            notes_conseiller    TEXT,
            chat_ia_synthese    TEXT,    -- Dernière réponse de l'IA

            -- Métadonnées
            date_creation       TEXT    NOT NULL,
            date_modification   TEXT    NOT NULL
        )
        """)

        # ----- Table des orientations EN ATTENTE -----
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS orientations_en_attente (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            -- Identité
            nom                 TEXT    NOT NULL,
            prenom              TEXT    NOT NULL,
            age                 INTEGER,
            sexe                TEXT,
            lycee               TEXT,
            choix_personnel     TEXT,
            projet_pro          TEXT,
            revenu_famille      TEXT,

            -- Orientation provisoire
            serie_provisoire    TEXT,              -- 'C' ou 'A' (proposition)
            statut              TEXT    NOT NULL,  -- 'attente' ou 'indetermine'
            score_confiance     INTEGER,
            trimestre_actuel    TEXT,              -- Trimestre des dernières notes

            -- Objectifs à atteindre
            objectif_moy_sci    REAL,              -- Moyenne scientifique cible
            objectif_moy_lit    REAL,              -- Moyenne littéraire cible

            -- Tests psychotechniques (bruts)
            d48_brut            REAL,
            krx_brut            REAL,
            meca_brut           REAL,
            bv11_brut           REAL,
            prc_brut            REAL,

            -- Aptitudes calculées
            SA_etal             REAL,
            LA_etal             REAL,

            -- Notes T1
            maths_t1            REAL,
            sci_phy_t1          REAL,
            svt_t1              REAL,
            francais_t1         REAL,
            histgeo_t1          REAL,
            anglais_t1          REAL,
            moy_sci_t1          REAL,
            moy_lit_t1          REAL,

            -- Notes T2 (si saisies)
            maths_t2            REAL,
            sci_phy_t2          REAL,
            svt_t2              REAL,
            francais_t2         REAL,
            histgeo_t2          REAL,
            anglais_t2          REAL,
            moy_sci_t2          REAL,
            moy_lit_t2          REAL,

            -- Enrichissement
            notes_conseiller    TEXT,
            chat_ia_synthese    TEXT,

            -- Alerte suivi
            prochain_rdv        TEXT,   -- Date de suivi recommandée (texte libre)
            alerte_envoyee      INTEGER DEFAULT 0,  -- 0=non, 1=oui

            -- Métadonnées
            date_creation       TEXT    NOT NULL,
            date_modification   TEXT    NOT NULL
        )
        """)

        # ----- Table des orientations EN PROBATION -----
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS orientations_probation (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            -- Identité
            nom                 TEXT    NOT NULL,
            prenom              TEXT    NOT NULL,
            age                 INTEGER,
            sexe                TEXT,
            lycee               TEXT,
            choix_personnel     TEXT,
            projet_pro          TEXT,
            revenu_famille      TEXT,

            -- Orientation sous condition
            serie_cible         TEXT,              -- Série visée si amélioration
            serie_risque        TEXT,              -- Série par défaut si échec
            statut              TEXT    NOT NULL DEFAULT 'probation',
            score_confiance     INTEGER,
            trimestre_actuel    TEXT,

            -- Objectifs IMPÉRATIFS
            objectif_moy_sci    REAL,
            objectif_moy_lit    REAL,
            deadline_trimestre  TEXT,              -- 'T3' — deadline de la décision finale

            -- Tests psychotechniques
            d48_brut            REAL,
            krx_brut            REAL,
            meca_brut           REAL,
            bv11_brut           REAL,
            prc_brut            REAL,

            -- Aptitudes
            SA_etal             REAL,
            LA_etal             REAL,

            -- Notes T1
            maths_t1            REAL,
            sci_phy_t1          REAL,
            svt_t1              REAL,
            francais_t1         REAL,
            histgeo_t1          REAL,
            anglais_t1          REAL,
            moy_sci_t1          REAL,
            moy_lit_t1          REAL,

            -- Notes T2
            maths_t2            REAL,
            sci_phy_t2          REAL,
            svt_t2              REAL,
            francais_t2         REAL,
            histgeo_t2          REAL,
            anglais_t2          REAL,
            moy_sci_t2          REAL,
            moy_lit_t2          REAL,

            -- Notes T3 (à saisir lors du suivi)
            maths_t3            REAL,
            sci_phy_t3          REAL,
            svt_t3              REAL,
            francais_t3         REAL,
            histgeo_t3          REAL,
            anglais_t3          REAL,
            moy_sci_t3          REAL,
            moy_lit_t3          REAL,

            -- Suivi
            nb_relances         INTEGER DEFAULT 0,   -- Nombre de fois relancé
            notes_conseiller    TEXT,
            chat_ia_synthese    TEXT,
            commentaire_suivi   TEXT,                -- Notes du conseiller lors du suivi

            -- Métadonnées
            date_creation       TEXT    NOT NULL,
            date_modification   TEXT    NOT NULL
        )
        """)

        # ----- Table de l'historique des changements de statut -----
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS historique_statuts (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            nom             TEXT    NOT NULL,
            prenom          TEXT    NOT NULL,
            lycee           TEXT,
            ancien_statut   TEXT,
            nouveau_statut  TEXT    NOT NULL,
            serie           TEXT,
            commentaire     TEXT,
            date_changement TEXT    NOT NULL
        )
        """)

        # ----- S7 — Table des brouillons (auto-save anti-perte de données) -----
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS brouillons_session (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id      TEXT    NOT NULL UNIQUE,   -- UUID 8 chars généré côté app
            step            INTEGER DEFAULT 0,
            nom             TEXT,
            prenom          TEXT,
            data_json       TEXT    NOT NULL,          -- Sérialisation complète du session_state
            timestamp       TEXT    NOT NULL
        )
        """)

        self.conn.commit()

    # -------------------------------------------------------------------------
    # S7 — BROUILLONS : AUTO-SAVE & REPRISE
    # -------------------------------------------------------------------------
    def sauvegarder_brouillon(self, session_id: str, step: int, data: dict):
        """Sauvegarde ou remplace le brouillon de la session courante.
        Appelé automatiquement à chaque transition d'étape.
        """
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            # Nettoyage des valeurs non-sérialisables
            clean = {}
            for k, v in data.items():
                try:
                    json.dumps(v)
                    clean[k] = v
                except (TypeError, ValueError):
                    clean[k] = str(v)

            self.conn.execute("""
                INSERT INTO brouillons_session
                    (session_id, step, nom, prenom, data_json, timestamp)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(session_id) DO UPDATE SET
                    step      = excluded.step,
                    nom       = excluded.nom,
                    prenom    = excluded.prenom,
                    data_json = excluded.data_json,
                    timestamp = excluded.timestamp
            """, (
                session_id,
                step,
                data.get("nom", ""),
                data.get("prenom", ""),
                json.dumps(clean),
                now,
            ))
            self.conn.commit()
        except Exception:
            pass  # L'auto-save ne doit jamais bloquer l'interface

    def charger_brouillon(self, session_id: str) -> Optional[dict]:
        """Charge un brouillon de session par son identifiant.
        Retourne le dict du session_state restauré, ou None si introuvable.
        """
        try:
            row = self.conn.execute(
                "SELECT data_json FROM brouillons_session WHERE session_id = ?",
                (session_id,)
            ).fetchone()
            if row:
                return json.loads(row["data_json"])
            return None
        except Exception:
            return None

    def supprimer_brouillon(self, session_id: str):
        """Supprime un brouillon après validation définitive du dossier."""
        try:
            self.conn.execute(
                "DELETE FROM brouillons_session WHERE session_id = ?",
                (session_id,)
            )
            self.conn.commit()
        except Exception:
            pass

    # -------------------------------------------------------------------------
    # SAUVEGARDE PRINCIPALE
    # -------------------------------------------------------------------------
    def sauvegarder_dossier(self, ss, etalonner_fn=None) -> dict:
        """
        Sauvegarde ou met à jour un dossier d'orientation depuis le session_state
        de Streamlit. Détermine automatiquement la bonne table selon le statut.

        Args:
            ss           : st.session_state (ou dict équivalent)
            etalonner_fn : Fonction d'étalonnage (optionnel, recalcule les scores)

        Returns:
            dict avec 'succes', 'table', 'id', 'message'

        Exemple :
            from capavenir_database import CapAvenirDB
            db = CapAvenirDB()
            resultat = db.sauvegarder_dossier(st.session_state, etalonner)
            if resultat['succes']:
                st.success(f"Dossier sauvegardé (ID {resultat['id']})")
        """
        statut = getattr(ss, "statut", None) or ss.get("statut", "confirme")
        if statut not in STATUTS_VALIDES:
            return {"succes": False, "message": f"Statut inconnu : {statut}"}

        table  = TABLE_PAR_STATUT[statut]
        now    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        def _get(key, default=None):
            """Lit une valeur depuis session_state (dict ou objet)."""
            if isinstance(ss, dict):
                return ss.get(key, default)
            return getattr(ss, key, default)

        nom    = (_get("nom", "") or "").strip().upper()
        prenom = (_get("prenom", "") or "").strip().capitalize()

        # ── S1 — Étalonnage par test_key ──
        # Tables internes minimales (pour que la DB soit autonome sans importer l'app)
        _ETAL = {
            "d48":  {0:2.0,2:4.0,4:6.0,6:8.0,8:9.5,10:11.0,12:12.5,14:13.5,
                     15:14.0,16:15.0,17:16.0,18:17.5,19:18.5,20:20.0},
            "krx":  {0:2.5,2:4.5,4:6.5,6:8.5,8:10.0,10:11.5,12:12.5,14:13.5,
                     15:14.5,16:15.5,17:16.5,18:17.5,19:19.0,20:20.0},
            "meca": {0:3.0,2:5.0,4:7.0,6:8.5,8:10.0,10:11.5,12:13.0,14:14.0,
                     15:14.5,16:15.5,17:16.5,18:17.5,19:18.5,20:20.0},
            "bv11": {0:3.5,2:5.5,4:7.5,6:9.0,8:10.5,10:12.0,12:13.0,14:14.0,
                     15:14.5,16:15.5,17:16.5,18:17.5,19:19.0,20:20.0},
            "prc":  {0:3.0,2:5.0,4:7.0,6:8.5,8:10.0,10:11.5,12:12.5,14:13.5,
                     15:14.5,16:15.5,17:16.5,18:17.5,19:18.5,20:20.0},
        }
        _ETAL_DEF = {0:3.0,1:4.5,2:5.5,3:6.5,4:7.5,5:8.5,6:9.5,7:10.0,
                     8:11.0,9:11.5,10:12.0,11:12.5,12:13.0,13:13.5,14:14.0,
                     15:14.5,16:15.5,17:16.5,18:17.5,19:18.5,20:20.0}

        def _etal_local(v, tkey):
            tbl = _ETAL.get(tkey, _ETAL_DEF)
            ks  = sorted(tbl.keys())
            bst = min(ks, key=lambda k: abs(k - v))
            return tbl[bst]

        def etal(v, tkey="default"):
            if etalonner_fn:
                try:
                    return etalonner_fn(float(v), tkey)
                except TypeError:
                    return etalonner_fn(float(v))
            return _etal_local(float(v), tkey)

        # ── Calcul des aptitudes ──
        d48  = float(_get("d48",  0.0) or 0.0)
        krx  = float(_get("krx",  0.0) or 0.0)
        meca = float(_get("meca", 0.0) or 0.0)
        bv11 = float(_get("bv11", 0.0) or 0.0)
        prc  = float(_get("prc",  0.0) or 0.0)

        SA_brut  = (krx + d48) / 2
        LA_brut  = (bv11 + prc) / 2
        SA_etal  = (etal(krx, "krx") + etal(d48, "d48")) / 2
        LA_etal  = (etal(bv11, "bv11") + etal(prc, "prc")) / 2
        MECA_etal = etal(meca, "meca")   # S2 — stocké pour stats filière technique

        # ── S3 — Coefficients pondérés camerounais ──
        COEFF_S = {"maths": 5, "sci_phy": 4, "svt": 2}
        COEFF_L = {"francais": 5, "histgeo": 3, "anglais": 2}

        # ── S5 — Notes scolaires : None sécurisé ──
        def notes_trim(t):
            """S5 — Gère les notes None (non saisies) sans produire de division par zéro."""
            def _n(k):
                v = _get(f"{k}_{t}")
                return float(v) if v is not None else None

            maths    = _n("maths")
            sci_phy  = _n("sci_phy")
            svt      = _n("svt")
            francais = _n("francais")
            histgeo  = _n("histgeo")
            anglais  = _n("anglais")

            # S3 — Moyenne pondérée (0.0 si note non renseignée)
            vals_sci = {m: _n(m) or 0.0 for m in ["maths","sci_phy","svt"]}
            vals_lit = {m: _n(m) or 0.0 for m in ["francais","histgeo","anglais"]}
            moy_sci = round(
                sum(vals_sci[m] * COEFF_S[m] for m in COEFF_S) / sum(COEFF_S.values()), 2
            )
            moy_lit = round(
                sum(vals_lit[m] * COEFF_L[m] for m in COEFF_L) / sum(COEFF_L.values()), 2
            )
            return {
                f"maths_{t}": maths, f"sci_phy_{t}": sci_phy, f"svt_{t}": svt,
                f"francais_{t}": francais, f"histgeo_{t}": histgeo, f"anglais_{t}": anglais,
                f"moy_sci_{t}": moy_sci, f"moy_lit_{t}": moy_lit,
            }

        n_t1 = notes_trim("t1")
        n_t2 = notes_trim("t2") if _get("t2_renseigne", False) else {k: None for k in notes_trim("t2")}
        n_t3 = notes_trim("t3") if _get("t3_renseigne", False) else {k: None for k in notes_trim("t3")}

        # ── Trimestre actif ──
        if _get("t3_renseigne", False):
            trim_actuel = "T3"
        elif _get("t2_renseigne", False):
            trim_actuel = "T2"
        else:
            trim_actuel = "T1"

        # ── Chat IA ──
        chat_history = _get("chat_history", [])
        ia_messages  = [m["content"] for m in chat_history if isinstance(m, dict) and m.get("role") == "ia"]
        ia_synthese  = ia_messages[-1] if ia_messages else ""

        # ── Objectifs (pour attente/probation) ──
        moy_sci_active = n_t1["moy_sci_t1"]
        moy_lit_active = n_t1["moy_lit_t1"]
        if trim_actuel == "T2":
            moy_sci_active = n_t2.get("moy_sci_t2") or moy_sci_active
            moy_lit_active = n_t2.get("moy_lit_t2") or moy_lit_active
        elif trim_actuel == "T3":
            moy_sci_active = n_t3.get("moy_sci_t3") or moy_sci_active
            moy_lit_active = n_t3.get("moy_lit_t3") or moy_lit_active

        obj_sci = round(min(20.0, moy_sci_active + 2.0), 1) if SA_etal > LA_etal else None
        obj_lit = round(min(20.0, moy_lit_active + 2.0), 1) if LA_etal > SA_etal else None

        # ── Orientation ──
        serie           = _get("orientation_finale", "?") or "?"
        score_confiance = int(_get("score_confiance", 0) or 0)
        notes_cons      = _get("notes_conseiller", "") or ""

        # ── Données communes à toutes les tables ──
        base = {
            "nom": nom, "prenom": prenom,
            "age": int(_get("age", 0) or 0),
            "sexe": _get("sexe", ""),
            "lycee": _get("lycee", "") or "",
            "choix_personnel": _get("choix_personnel", "") or "",
            "projet_pro": _get("projet_pro", "") or "",
            "revenu_famille": _get("revenu", "") or "",
            "score_confiance": score_confiance,
            "trimestre_actuel": trim_actuel,
            "d48_brut": d48, "krx_brut": krx, "meca_brut": meca,
            "bv11_brut": bv11, "prc_brut": prc,
            "SA_etal": SA_etal, "LA_etal": LA_etal,
            # S2 — MECA étalonnée exposée pour les stats filière technique
            "meca_etal": MECA_etal,
            **n_t1, **n_t2, **n_t3,
            "notes_conseiller": notes_cons,
            "chat_ia_synthese": ia_synthese,
            "date_modification": now,
        }

        # ── Données supplémentaires selon la table ──
        if table == "orientations_confirmees":
            data = {
                **base,
                "serie_finale": serie,
                "statut": statut,
                "trimestre_decision": trim_actuel,
                # S1 — étalonnage par test
                "d48_etal":  etal(d48,  "d48"),
                "krx_etal":  etal(krx,  "krx"),
                "meca_etal": etal(meca, "meca"),
                "bv11_etal": etal(bv11, "bv11"),
                "prc_etal":  etal(prc,  "prc"),
                "SA_brut": SA_brut, "LA_brut": LA_brut,
            }

        elif table == "orientations_en_attente":
            data = {
                **base,
                "serie_provisoire": serie,
                "statut": statut,
                "objectif_moy_sci": obj_sci,
                "objectif_moy_lit": obj_lit,
                "prochain_rdv": f"Saisie notes {('T2' if trim_actuel == 'T1' else 'T3')}",
            }

        elif table == "orientations_probation":
            data = {
                **base,
                # S2 — TECHNIQUE stocké dans serie_cible
                "serie_cible": serie,
                "serie_risque": "A" if serie == "C" else ("C" if serie == "A" else "TECHNIQUE"),
                "statut": "probation" if serie not in ("TECHNIQUE",) else "confirme",
                "objectif_moy_sci": obj_sci,
                "objectif_moy_lit": obj_lit,
                "deadline_trimestre": "T3",
                "SA_brut": SA_brut, "LA_brut": LA_brut,
            }
        else:
            return {"succes": False, "message": f"Table non reconnue : {table}"}

        # ── Supprimer l'ancien enregistrement si le statut a changé de table ──
        lycee_val = _get("lycee", "") or ""
        ancien_statut_autre_table = self._supprimer_de_autres_tables(
            nom, prenom, lycee_val, table
        )

        # ── Vérifier si le dossier existe déjà dans la table CIBLE ──
        existant = self._trouver_existant(nom, prenom, lycee_val, table)

        try:
            if existant:
                dossier_id = existant["id"]
                self._update(table, data, dossier_id)
                action = "mis à jour"
            else:
                data["date_creation"] = now
                dossier_id = self._insert(table, data)
                action = "créé"

            # Déterminer l'ancien statut pour le log
            ancien_pour_log = (existant["statut"] if existant
                               else ancien_statut_autre_table)

            # ── Log dans l'historique ──
            self._log_historique(
                nom=nom, prenom=prenom,
                lycee=lycee_val,
                ancien_statut=ancien_pour_log,
                nouveau_statut=statut,
                serie=serie,
                commentaire=f"Dossier {action} — trimestre {trim_actuel}"
                            + (f" (migration depuis {ancien_pour_log})"
                               if ancien_statut_autre_table else ""),
            )

            return {
                "succes": True,
                "table": table,
                "id": dossier_id,
                "action": action,
                "message": f"✅ Dossier de {prenom} {nom} {action} (ID {dossier_id})"
            }

        except sqlite3.Error as e:
            return {"succes": False, "message": f"Erreur SQLite : {str(e)}"}

    # -------------------------------------------------------------------------
    # LECTURE / REQUÊTES
    # -------------------------------------------------------------------------
    def lister_dossiers(self, statut: str, lycee: str = None) -> list:
        """
        Retourne tous les dossiers d'un statut donné.

        Args:
            statut : 'confirme', 'revise', 'attente', 'probation', 'indetermine'
            lycee  : Filtre optionnel par lycée

        Returns:
            Liste de dicts (chaque dict = un dossier)

        Exemple :
            dossiers = db.lister_dossiers("probation")
            for d in dossiers:
                print(d["nom"], d["prenom"], d["serie_cible"])
        """
        table = TABLE_PAR_STATUT.get(statut)
        if not table:
            return []

        if statut in ("confirme", "revise"):
            where_statut = f"statut = '{statut}'"
        else:
            where_statut = "1=1"  # Toute la table

        query = f"SELECT * FROM {table} WHERE {where_statut}"
        params = []
        if lycee:
            query += " AND lycee = ?"
            params.append(lycee)
        query += " ORDER BY UPPER(nom) ASC, UPPER(prenom) ASC"

        try:
            rows = self.conn.execute(query, params).fetchall()
            return [dict(row) for row in rows]
        except sqlite3.Error:
            return []

    def rechercher_eleve(self, nom: str, prenom: str = None, lycee: str = None) -> list:
        """
        Recherche un élève dans toutes les tables.

        Args:
            nom    : Nom de famille (insensible à la casse)
            prenom : Prénom (optionnel)
            lycee  : Lycée (optionnel)

        Returns:
            Liste de dossiers trouvés (avec champ 'table_source')

        Exemple :
            resultats = db.rechercher_eleve("MBALLA")
            resultats = db.rechercher_eleve("MBALLA", "Jean", "Lycée de Douala")
        """
        tables = [
            ("orientations_confirmees", "confirme/revise"),
            ("orientations_en_attente", "attente"),
            ("orientations_probation",  "probation"),
        ]
        resultats = []
        for table, label in tables:
            query  = f"SELECT * FROM {table} WHERE UPPER(nom) = UPPER(?)"
            params = [nom.strip()]
            if prenom:
                query  += " AND UPPER(prenom) = UPPER(?)"
                params.append(prenom.strip())
            if lycee:
                query  += " AND lycee = ?"
                params.append(lycee)
            try:
                rows = self.conn.execute(query, params).fetchall()
                for row in rows:
                    d = dict(row)
                    d["table_source"] = label
                    resultats.append(d)
            except sqlite3.Error:
                continue
        return resultats

    def get_dossier_par_id(self, dossier_id: int, statut: str) -> Optional[dict]:
        """
        Récupère un dossier précis par son ID et son statut.

        Args:
            dossier_id : L'identifiant numérique du dossier
            statut     : Le statut pour savoir dans quelle table chercher

        Returns:
            dict du dossier, ou None si non trouvé
        """
        table = TABLE_PAR_STATUT.get(statut)
        if not table:
            return None
        try:
            row = self.conn.execute(
                f"SELECT * FROM {table} WHERE id = ?", (dossier_id,)
            ).fetchone()
            return dict(row) if row else None
        except sqlite3.Error:
            return None

    def statistiques(self) -> dict:
        """
        Retourne des statistiques globales sur la base de données.

        Returns:
            dict avec total, nb_confirmes, nb_attente, nb_probation,
            nb_revises, nb_serie_c, nb_serie_a, repartition_lycees

        Exemple :
            stats = db.statistiques()
            print(f"Total dossiers : {stats['total']}")
        """
        stats = {
            "nb_confirmes":   0,
            "nb_revises":     0,
            "nb_attente":     0,
            "nb_indetermines":0,
            "nb_probation":   0,
            "nb_serie_c":     0,
            "nb_serie_a":     0,
            "total":          0,
            "repartition_lycees": {},
        }
        try:
            stats["nb_confirmes"]   = self.conn.execute(
                "SELECT COUNT(*) FROM orientations_confirmees WHERE statut='confirme'"
            ).fetchone()[0]
            stats["nb_revises"]     = self.conn.execute(
                "SELECT COUNT(*) FROM orientations_confirmees WHERE statut='revise'"
            ).fetchone()[0]
            stats["nb_attente"]     = self.conn.execute(
                "SELECT COUNT(*) FROM orientations_en_attente WHERE statut='attente'"
            ).fetchone()[0]
            stats["nb_indetermines"]= self.conn.execute(
                "SELECT COUNT(*) FROM orientations_en_attente WHERE statut='indetermine'"
            ).fetchone()[0]
            stats["nb_probation"]   = self.conn.execute(
                "SELECT COUNT(*) FROM orientations_probation"
            ).fetchone()[0]

            # Série C toutes tables
            stats["nb_serie_c"] = (
                self.conn.execute(
                    "SELECT COUNT(*) FROM orientations_confirmees WHERE serie_finale='C'"
                ).fetchone()[0] +
                self.conn.execute(
                    "SELECT COUNT(*) FROM orientations_en_attente WHERE serie_provisoire='C'"
                ).fetchone()[0] +
                self.conn.execute(
                    "SELECT COUNT(*) FROM orientations_probation WHERE serie_cible='C'"
                ).fetchone()[0]
            )

            # Série A toutes tables
            stats["nb_serie_a"] = (
                self.conn.execute(
                    "SELECT COUNT(*) FROM orientations_confirmees WHERE serie_finale='A'"
                ).fetchone()[0] +
                self.conn.execute(
                    "SELECT COUNT(*) FROM orientations_en_attente WHERE serie_provisoire='A'"
                ).fetchone()[0] +
                self.conn.execute(
                    "SELECT COUNT(*) FROM orientations_probation WHERE serie_cible='A'"
                ).fetchone()[0]
            )

            # S2 — Filière TECHNIQUE
            stats["nb_serie_technique"] = (
                self.conn.execute(
                    "SELECT COUNT(*) FROM orientations_confirmees WHERE serie_finale='TECHNIQUE'"
                ).fetchone()[0] +
                self.conn.execute(
                    "SELECT COUNT(*) FROM orientations_probation WHERE serie_cible='TECHNIQUE'"
                ).fetchone()[0]
            )

            stats["total"] = (
                stats["nb_confirmes"] + stats["nb_revises"] +
                stats["nb_attente"] + stats["nb_indetermines"] +
                stats["nb_probation"]
            )

            # Répartition par lycée (table confirmées)
            rows = self.conn.execute(
                "SELECT lycee, COUNT(*) as n FROM orientations_confirmees GROUP BY lycee"
            ).fetchall()
            stats["repartition_lycees"] = {r["lycee"]: r["n"] for r in rows}

        except sqlite3.Error:
            pass
        return stats

    def historique_eleve(self, nom: str, prenom: str) -> list:
        """
        Retourne l'historique complet des changements de statut d'un élève.

        Exemple :
            hist = db.historique_eleve("MBALLA", "Jean")
        """
        try:
            rows = self.conn.execute(
                """SELECT * FROM historique_statuts
                   WHERE UPPER(nom) = UPPER(?) AND UPPER(prenom) = UPPER(?)
                   ORDER BY date_changement DESC""",
                (nom.strip(), prenom.strip())
            ).fetchall()
            return [dict(r) for r in rows]
        except sqlite3.Error:
            return []

    # -------------------------------------------------------------------------
    # MISE À JOUR CIBLÉE
    # -------------------------------------------------------------------------
    def promouvoir_dossier(self, dossier_id: int,
                           ancien_statut: str, nouveau_statut: str,
                           commentaire: str = "") -> dict:
        """
        Déplace un dossier d'une table à l'autre suite à un changement de statut.
        Ex : passer de 'attente' → 'probation' ou 'probation' → 'confirme'.

        Args:
            dossier_id    : ID du dossier dans l'ancienne table
            ancien_statut : Statut actuel ('attente', 'probation', etc.)
            nouveau_statut: Nouveau statut cible
            commentaire   : Note du conseiller

        Returns:
            dict avec 'succes', 'nouvel_id', 'message'

        Exemple :
            res = db.promouvoir_dossier(12, "attente", "probation", "Pas d'amélioration au T2")
            if res['succes']:
                st.success(res['message'])
        """
        ancienne_table = TABLE_PAR_STATUT.get(ancien_statut)
        nouvelle_table = TABLE_PAR_STATUT.get(nouveau_statut)
        if not ancienne_table or not nouvelle_table:
            return {"succes": False, "message": "Statut non reconnu."}
        if ancienne_table == nouvelle_table:
            # Même table : simple mise à jour du statut
            try:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.conn.execute(
                    f"UPDATE {ancienne_table} SET statut=?, date_modification=? WHERE id=?",
                    (nouveau_statut, now, dossier_id)
                )
                self.conn.commit()
                return {"succes": True, "nouvel_id": dossier_id,
                        "message": f"Statut mis à jour → {nouveau_statut}"}
            except sqlite3.Error as e:
                return {"succes": False, "message": str(e)}

        # Tables différentes : copier puis supprimer
        try:
            row = self.conn.execute(
                f"SELECT * FROM {ancienne_table} WHERE id = ?", (dossier_id,)
            ).fetchone()
            if not row:
                return {"succes": False, "message": f"Dossier {dossier_id} introuvable."}

            d = dict(row)
            d.pop("id", None)
            d["statut"]           = nouveau_statut
            d["date_modification"]= datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if commentaire:
                d["notes_conseiller"] = (d.get("notes_conseiller") or "") + f"\n[{d['date_modification']}] {commentaire}"

            nouvel_id = self._insert(nouvelle_table, d)
            self.conn.execute(f"DELETE FROM {ancienne_table} WHERE id = ?", (dossier_id,))
            self.conn.commit()

            self._log_historique(
                nom=d["nom"], prenom=d["prenom"], lycee=d.get("lycee", ""),
                ancien_statut=ancien_statut, nouveau_statut=nouveau_statut,
                serie=d.get("serie_finale") or d.get("serie_provisoire") or d.get("serie_cible", "?"),
                commentaire=commentaire or f"Promotion {ancien_statut} → {nouveau_statut}",
            )
            return {"succes": True, "nouvel_id": nouvel_id,
                    "message": f"✅ Dossier déplacé vers '{nouvelle_table}' (ID {nouvel_id})"}

        except sqlite3.Error as e:
            self.conn.rollback()
            return {"succes": False, "message": str(e)}

    def mettre_a_jour_notes_t3(self, dossier_id: int,
                                notes_t3: dict, commentaire: str = "") -> dict:
        """
        Met à jour les notes T3 d'un dossier en probation.

        Args:
            dossier_id : ID dans orientations_probation
            notes_t3   : dict avec clés maths_t3, sci_phy_t3, svt_t3,
                         francais_t3, histgeo_t3, anglais_t3
            commentaire: Note du conseiller

        Returns:
            dict avec 'succes', 'message', 'moy_sci_t3', 'moy_lit_t3'

        Exemple :
            db.mettre_a_jour_notes_t3(5, {
                "maths_t3": 13.0, "sci_phy_t3": 12.5, "svt_t3": 11.0,
                "francais_t3": 10.0, "histgeo_t3": 11.0, "anglais_t3": 12.0
            }, "Notes T3 saisies lors de la visite du 15/06")
        """
        moy_sci_t3 = round((notes_t3.get("maths_t3", 0) +
                            notes_t3.get("sci_phy_t3", 0) +
                            notes_t3.get("svt_t3", 0)) / 3, 2)
        moy_lit_t3 = round((notes_t3.get("francais_t3", 0) +
                            notes_t3.get("histgeo_t3", 0) +
                            notes_t3.get("anglais_t3", 0)) / 3, 2)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            self.conn.execute("""
                UPDATE orientations_probation
                SET maths_t3=?, sci_phy_t3=?, svt_t3=?,
                    francais_t3=?, histgeo_t3=?, anglais_t3=?,
                    moy_sci_t3=?, moy_lit_t3=?,
                    commentaire_suivi=?, date_modification=?,
                    nb_relances = nb_relances + 1
                WHERE id = ?
            """, (
                notes_t3.get("maths_t3"), notes_t3.get("sci_phy_t3"), notes_t3.get("svt_t3"),
                notes_t3.get("francais_t3"), notes_t3.get("histgeo_t3"), notes_t3.get("anglais_t3"),
                moy_sci_t3, moy_lit_t3,
                commentaire, now,
                dossier_id
            ))
            self.conn.commit()
            return {
                "succes": True,
                "moy_sci_t3": moy_sci_t3,
                "moy_lit_t3": moy_lit_t3,
                "message": f"Notes T3 enregistrées. Moy. Sci. T3 = {moy_sci_t3:.2f}/20"
            }
        except sqlite3.Error as e:
            return {"succes": False, "message": str(e)}

    # -------------------------------------------------------------------------
    # SUPPRESSION
    # -------------------------------------------------------------------------
    def supprimer_dossier(self, dossier_id: int, statut: str) -> dict:
        """
        Supprime définitivement un dossier.

        Args:
            dossier_id : ID du dossier
            statut     : Statut pour localiser la bonne table

        Returns:
            dict avec 'succes', 'message'

        Exemple :
            db.supprimer_dossier(7, "attente")
        """
        table = TABLE_PAR_STATUT.get(statut)
        if not table:
            return {"succes": False, "message": "Statut non reconnu."}
        try:
            self.conn.execute(f"DELETE FROM {table} WHERE id = ?", (dossier_id,))
            self.conn.commit()
            return {"succes": True, "message": f"Dossier {dossier_id} supprimé."}
        except sqlite3.Error as e:
            return {"succes": False, "message": str(e)}

    # -------------------------------------------------------------------------
    # EXPORT
    # -------------------------------------------------------------------------
    def exporter_json(self, statut: str = None) -> str:
        """
        Exporte tous les dossiers (ou ceux d'un statut) en JSON.

        Args:
            statut : Si précisé, exporte uniquement ce statut

        Returns:
            Chaîne JSON

        Exemple :
            json_str = db.exporter_json("confirme")
            st.download_button("Export JSON", json_str, "export.json")
        """
        if statut:
            data = {statut: self.lister_dossiers(statut)}
        else:
            data = {
                "confirmes":    self.lister_dossiers("confirme"),
                "revises":      self.lister_dossiers("revise"),
                "en_attente":   self.lister_dossiers("attente"),
                "indetermines": self.lister_dossiers("indetermine"),
                "probation":    self.lister_dossiers("probation"),
                "statistiques": self.statistiques(),
                "export_date":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        return json.dumps(data, ensure_ascii=False, indent=2)

    # -------------------------------------------------------------------------
    # FERMETURE
    # -------------------------------------------------------------------------
    def close(self):
        """Ferme proprement la connexion à la base de données."""
        if self.conn:
            self.conn.close()

    # -------------------------------------------------------------------------
    # MÉTHODES INTERNES (privées)
    # -------------------------------------------------------------------------
    def _insert(self, table: str, data: dict) -> int:
        """Insère un dict dans la table et retourne l'ID généré."""
        # Garder uniquement les colonnes qui existent dans la table
        colonnes_valides = self._colonnes(table)
        data_filtre = {k: v for k, v in data.items() if k in colonnes_valides}

        cols   = ", ".join(data_filtre.keys())
        plages = ", ".join(["?"] * len(data_filtre))
        vals   = list(data_filtre.values())
        cursor = self.conn.execute(
            f"INSERT INTO {table} ({cols}) VALUES ({plages})", vals
        )
        self.conn.commit()
        return cursor.lastrowid

    def _update(self, table: str, data: dict, record_id: int):
        """Met à jour un enregistrement existant."""
        colonnes_valides = self._colonnes(table)
        data_filtre = {k: v for k, v in data.items()
                       if k in colonnes_valides and k not in ("id", "date_creation")}
        set_clause = ", ".join([f"{k} = ?" for k in data_filtre.keys()])
        vals = list(data_filtre.values()) + [record_id]
        self.conn.execute(f"UPDATE {table} SET {set_clause} WHERE id = ?", vals)
        self.conn.commit()

    def _colonnes(self, table: str) -> set:
        """Retourne les noms de colonnes d'une table."""
        rows = self.conn.execute(f"PRAGMA table_info({table})").fetchall()
        return {r["name"] for r in rows}

    def _supprimer_de_autres_tables(self, nom: str, prenom: str,
                                     lycee: str, table_cible: str) -> Optional[str]:
        """
        Cherche le dossier (nom/prenom/lycee) dans toutes les tables SAUF table_cible
        et le supprime si trouvé. Retourne l'ancien statut ou None.
        Évite les doublons quand le statut change de table.
        """
        TOUTES_TABLES = [
            ("orientations_confirmees", "confirme"),
            ("orientations_en_attente", "attente"),
            ("orientations_probation",  "probation"),
        ]
        ancien_statut = None
        for table, _ in TOUTES_TABLES:
            if table == table_cible:
                continue
            try:
                row = self.conn.execute(
                    f"""SELECT id, statut FROM {table}
                        WHERE UPPER(nom) = UPPER(?) AND UPPER(prenom) = UPPER(?)
                        AND lycee = ?""",
                    (nom, prenom, lycee)
                ).fetchone()
                if row:
                    ancien_statut = row["statut"]
                    self.conn.execute(
                        f"DELETE FROM {table} WHERE id = ?", (row["id"],)
                    )
                    self.conn.commit()
            except Exception:
                pass
        return ancien_statut

    def _trouver_existant(self, nom: str, prenom: str,
                          lycee: str, table: str) -> Optional[dict]:
        """Vérifie si un dossier existe déjà dans la table."""
        try:
            row = self.conn.execute(
                f"""SELECT id, statut FROM {table}
                    WHERE UPPER(nom) = UPPER(?) AND UPPER(prenom) = UPPER(?)
                    AND lycee = ?""",
                (nom, prenom, lycee)
            ).fetchone()
            return dict(row) if row else None
        except sqlite3.Error:
            return None

    def _log_historique(self, nom: str, prenom: str, lycee: str,
                        ancien_statut: Optional[str], nouveau_statut: str,
                        serie: str, commentaire: str = ""):
        """Enregistre un changement de statut dans l'historique."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            self.conn.execute("""
                INSERT INTO historique_statuts
                    (nom, prenom, lycee, ancien_statut, nouveau_statut,
                     serie, commentaire, date_changement)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (nom, prenom, lycee, ancien_statut, nouveau_statut,
                  serie, commentaire, now))
            self.conn.commit()
        except sqlite3.Error:
            pass  # L'historique ne doit pas bloquer l'opération principale


# =============================================================================
# SECTION D'INTÉGRATION — À COPIER dans streamlit_run_app.py
# =============================================================================
"""
────────────────────────────────────────────────────────────────────
GUIDE D'INTÉGRATION DANS streamlit_run_app.py
────────────────────────────────────────────────────────────────────

1. IMPORT (en haut du fichier, après les autres imports) :
────────────────────────────────────────────────────────
    from capavenir_database import CapAvenirDB

2. INITIALISATION (après st.set_page_config) :
──────────────────────────────────────────────
    @st.cache_resource
    def get_db():
        return CapAvenirDB()
    db = get_db()

3. SAUVEGARDE (dans l'étape 4 — Fiche finale, bouton "Valider") :
─────────────────────────────────────────────────────────────────
    if st.button("💾 Sauvegarder le dossier", type="primary"):
        resultat = db.sauvegarder_dossier(st.session_state, etalonner)
        if resultat["succes"]:
            st.success(resultat["message"])
        else:
            st.error(resultat["message"])

4. PAGE DE CONSULTATION (mode conseiller) :
───────────────────────────────────────────
    if mode_conseiller:
        st.sidebar.subheader("📂 Base de données")
        onglet = st.sidebar.selectbox("Table", ["confirme","attente","probation"])
        dossiers = db.lister_dossiers(onglet)
        st.sidebar.metric("Dossiers", len(dossiers))
        for d in dossiers:
            st.write(f"{d['nom']} {d['prenom']} — {d.get('serie_finale') or d.get('serie_provisoire', '?')}")

5. STATISTIQUES :
─────────────────
    stats = db.statistiques()
    col1.metric("✅ Confirmés",  stats["nb_confirmes"])
    col2.metric("⏳ En attente", stats["nb_attente"])
    col3.metric("⚠️ Probation",  stats["nb_probation"])
    col4.metric("🔄 Révisés",    stats["nb_revises"])

6. EXPORT JSON (bouton téléchargement) :
────────────────────────────────────────
    st.download_button(
        "📦 Exporter tous les dossiers (JSON)",
        db.exporter_json(),
        file_name="capavenir_export.json",
        mime="application/json"
    )

7. PROMOUVOIR un dossier (ex: attente → probation) :
────────────────────────────────────────────────────
    res = db.promouvoir_dossier(dossier_id, "attente", "probation",
                                 "Aucune progression au T2")
    if res["succes"]:
        st.success(res["message"])

────────────────────────────────────────────────────────────────────
"""


# =============================================================================
# TEST RAPIDE — Exécuter ce fichier directement pour vérifier la BDD
# python capavenir_database.py
# =============================================================================
if __name__ == "__main__":
    print("=" * 60)
    print("  CapAvenir CMR — Test de la base de données")
    print("=" * 60)

    db = CapAvenirDB("test_capavenir.db")

    # Données de test simulant un st.session_state
    test_data = {
        "nom": "MBALLA", "prenom": "Jean", "age": 16,
        "sexe": "Masculin", "lycee": "Lycée de Douala",
        "choix_personnel": "C (Scientifique)",
        "projet_pro": "Ingénieur informatique",
        "revenu": "Moyen (50 000 - 150 000 FCFA/mois)",
        "d48": 15.0, "krx": 14.0, "meca": 13.0, "bv11": 10.0, "prc": 9.5,
        "maths_t1": 14.0, "sci_phy_t1": 13.5, "svt_t1": 12.0,
        "francais_t1": 11.0, "histgeo_t1": 10.5, "anglais_t1": 12.0,
        "t2_renseigne": False, "t3_renseigne": False,
        "orientation_finale": "C",
        "statut": "confirme",
        "score_confiance": 85,
        "notes_conseiller": "Élève sérieux, profil scientifique solide.",
        "chat_history": [{"role": "ia", "content": "Profil C confirmé."}],
    }

    res = db.sauvegarder_dossier(test_data)
    print(f"\n[Sauvegarde 1] {res['message']}")

    # Test dossier en attente
    test_data2 = {**test_data, "nom": "FOKO", "prenom": "Marie",
                  "statut": "attente", "orientation_finale": "C",
                  "maths_t1": 8.0, "sci_phy_t1": 7.5, "svt_t1": 9.0,
                  "d48": 14.0, "krx": 13.0, "score_confiance": 40}
    res2 = db.sauvegarder_dossier(test_data2)
    print(f"[Sauvegarde 2] {res2['message']}")

    # Test dossier en probation
    test_data3 = {**test_data, "nom": "NKEMELI", "prenom": "Paul",
                  "statut": "probation", "orientation_finale": "A",
                  "t2_renseigne": True,
                  "maths_t2": 8.5, "sci_phy_t2": 7.0, "svt_t2": 9.0,
                  "francais_t2": 9.0, "histgeo_t2": 8.5, "anglais_t2": 10.0,
                  "d48": 8.0, "krx": 7.5, "bv11": 13.0, "prc": 14.0,
                  "score_confiance": 35}
    res3 = db.sauvegarder_dossier(test_data3)
    print(f"[Sauvegarde 3] {res3['message']}")

    # Statistiques
    stats = db.statistiques()
    print(f"\n[Stats] Total : {stats['total']} dossiers")
    print(f"  ✅ Confirmés : {stats['nb_confirmes']}")
    print(f"  ⏳ En attente: {stats['nb_attente']}")
    print(f"  ⚠️  Probation : {stats['nb_probation']}")
    print(f"  Série C      : {stats['nb_serie_c']} | Série A : {stats['nb_serie_a']}")

    # Recherche
    found = db.rechercher_eleve("MBALLA")
    print(f"\n[Recherche 'MBALLA'] {len(found)} résultat(s)")

    # Historique
    hist = db.historique_eleve("MBALLA", "Jean")
    print(f"[Historique MBALLA Jean] {len(hist)} entrée(s)")
    for h in hist:
        print(f"  → {h['ancien_statut']} → {h['nouveau_statut']} le {h['date_changement']}")

    db.close()

    # Nettoyage du fichier de test
    if os.path.exists("test_capavenir.db"):
        os.remove("test_capavenir.db")
        print("\n[Test terminé] Fichier test supprimé.")
    print("\n✅ Module capavenir_database.py opérationnel.")

