import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Search,
  Users,
  FileText,
  BookOpen,
  Plus,
  X,
  Copy,
  Check,
  ChevronDown,
  ChevronRight,
  AlertTriangle,
  Trash2,
  Clock,
  Scale,
  Upload,
  CalendarDays,
  LogOut,
  Lock,
  Printer,
  Download,
  Settings,
} from "lucide-react";
import * as XLSX from "xlsx";

/* ============================================================
   DONNÉES DE RÉFÉRENCE
   Base générique conforme aux grands principes du droit du
   travail français. À PERSONNALISER avec les articles réels
   du règlement intérieur 3 Media dès qu'il sera transmis.
   ============================================================ */

const CATEGORIES = [
  { id: "assiduite", label: "Assiduité & ponctualité" },
  { id: "qualite", label: "Qualité & conformité du travail" },
  { id: "comportement", label: "Comportement & discipline" },
  { id: "informatique", label: "Usage des outils informatiques" },
  { id: "confidentialite", label: "Confidentialité & sécurité des données" },
  { id: "malhonnetete", label: "Fraude & malhonnêteté" },
];

const GRAVITE = {
  leger: { label: "Léger", order: 1 },
  serieux: { label: "Sérieux", order: 2 },
  grave: { label: "Grave", order: 3 },
  lourd: { label: "Lourd", order: 4 },
};

const MANQUEMENTS = [
  { id: "m01", categorie: "assiduite", libelle: "Retard répété non justifié", description: "Retards récurrents au poste ou en pause, sans motif valable communiqué au superviseur.", gravite: "leger", sanction: "Rappel oral puis avertissement écrit en cas de récidive", article: "RI – Art. Assiduité (à compléter)", baseLegale: "Pouvoir disciplinaire de l'employeur – C. trav. art. L1331-1" },
  { id: "m02", categorie: "assiduite", libelle: "Absence injustifiée", description: "Absence sans justificatif transmis dans le délai prévu (arrêt de travail, motif personnel non prévenu).", gravite: "serieux", sanction: "Avertissement écrit, mise à pied en cas de répétition", article: "RI – Art. Assiduité (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m03", categorie: "assiduite", libelle: "Abandon de poste", description: "Départ du poste de travail sans autorisation pendant les horaires prévus, sans reprise de contact.", gravite: "grave", sanction: "Convocation à entretien préalable, mise à pied conservatoire possible", article: "RI – Art. Assiduité (à compléter)", baseLegale: "C. trav. art. L1237-1-1" },
  { id: "m04", categorie: "qualite", libelle: "Non-respect des procédures / scripts d'appel", description: "Écarts répétés avec les process qualité définis (script, argumentaire, étapes obligatoires de traitement).", gravite: "leger", sanction: "Recadrage oral, accompagnement, avertissement si récidive", article: "RI – Art. Qualité (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m05", categorie: "qualite", libelle: "Erreurs de traitement client répétées", description: "Erreurs récurrentes malgré accompagnement (mauvaise saisie, information erronée transmise au client).", gravite: "serieux", sanction: "Avertissement écrit, plan d'accompagnement formalisé", article: "RI – Art. Qualité (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m06", categorie: "qualite", libelle: "Non-atteinte prolongée des objectifs malgré accompagnement", description: "Sous-performance persistante après plans d'action successifs et entretiens de suivi documentés.", gravite: "serieux", sanction: "Avertissement écrit, évaluation formelle", article: "RI – Art. Qualité (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m07", categorie: "comportement", libelle: "Insubordination", description: "Refus caractérisé d'exécuter une consigne légitime du superviseur ou du management.", gravite: "grave", sanction: "Convocation à entretien préalable", article: "RI – Art. Discipline (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m08", categorie: "comportement", libelle: "Propos injurieux ou irrespectueux envers un collègue", description: "Comportement verbal inapproprié, agressivité ou manque de respect envers un membre de l'équipe.", gravite: "grave", sanction: "Convocation à entretien préalable, avertissement selon gravité", article: "RI – Art. Discipline (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m09", categorie: "comportement", libelle: "Propos injurieux envers un client", description: "Manque de respect, agressivité ou propos déplacés tenus envers un client durant un appel.", gravite: "grave", sanction: "Convocation à entretien préalable, mise à pied possible", article: "RI – Art. Discipline (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m10", categorie: "comportement", libelle: "Situation de harcèlement ou discrimination", description: "Comportement répété portant atteinte à la dignité d'un collaborateur (moral, sexuel, discriminatoire).", gravite: "lourd", sanction: "Convocation immédiate, mise à pied conservatoire, procédure renforcée", article: "RI – Art. Discipline (à compléter)", baseLegale: "C. trav. art. L1152-1 / L1153-1" },
  { id: "m11", categorie: "informatique", libelle: "Usage abusif d'internet / réseaux sociaux sur le temps de travail", description: "Utilisation manifestement excessive et non professionnelle des outils numériques pendant les heures de travail.", gravite: "leger", sanction: "Rappel oral, avertissement en cas de récidive", article: "Charte informatique – Art. (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m12", categorie: "informatique", libelle: "Non-respect de la charte informatique", description: "Installation de logiciels non autorisés, contournement des règles de sécurité du SI.", gravite: "serieux", sanction: "Avertissement écrit", article: "Charte informatique – Art. (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m13", categorie: "informatique", libelle: "Consultation de contenus interdits", description: "Consultation de sites ou contenus non conformes à la charte informatique (illégaux, inappropriés).", gravite: "grave", sanction: "Convocation à entretien préalable", article: "Charte informatique – Art. (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m14", categorie: "confidentialite", libelle: "Divulgation d'informations client à un tiers non habilité", description: "Communication d'informations personnelles ou contractuelles client en dehors du cadre autorisé.", gravite: "grave", sanction: "Convocation à entretien préalable, mise à pied conservatoire possible", article: "RI – Art. Confidentialité (à compléter)", baseLegale: "RGPD art. 5 / C. trav. art. L1331-1" },
  { id: "m15", categorie: "confidentialite", libelle: "Non-respect des procédures RGPD", description: "Manquement aux règles de traitement des données personnelles (accès non justifié, absence de purge, etc.).", gravite: "serieux", sanction: "Avertissement écrit, rappel des obligations RGPD", article: "RI – Art. Confidentialité (à compléter)", baseLegale: "RGPD art. 5" },
  { id: "m16", categorie: "confidentialite", libelle: "Violation du secret professionnel", description: "Divulgation d'informations internes sensibles (process, données EDF, informations stratégiques).", gravite: "lourd", sanction: "Convocation immédiate, procédure renforcée", article: "RI – Art. Confidentialité (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m17", categorie: "malhonnetete", libelle: "Falsification de données ou de reporting", description: "Modification volontaire de statistiques, de comptes rendus ou de données de production.", gravite: "grave", sanction: "Convocation à entretien préalable, mise à pied conservatoire possible", article: "RI – Art. Discipline (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m18", categorie: "malhonnetete", libelle: "Vol ou détournement", description: "Appropriation de biens de l'entreprise, de collègues ou de clients.", gravite: "lourd", sanction: "Convocation immédiate, mise à pied conservatoire, procédure de licenciement", article: "RI – Art. Discipline (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m19", categorie: "malhonnetete", libelle: "Fraude (badgeage, primes, indicateurs)", description: "Manipulation intentionnelle des systèmes de badgeage, de suivi d'activité ou de calcul de primes.", gravite: "lourd", sanction: "Convocation immédiate, procédure de licenciement", article: "RI – Art. Discipline (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m20", categorie: "assiduite", libelle: "Dépassement du temps de pause", description: "Durée de pause réellement prise supérieure à la durée autorisée, de façon répétée ou significative.", gravite: "leger", sanction: "Rappel oral puis avertissement écrit en cas de récidive", article: "RI – Art. Assiduité (à compléter)", baseLegale: "C. trav. art. L1331-1" },
  { id: "m21", categorie: "assiduite", libelle: "Départ anticipé non autorisé", description: "Départ du poste avant l'heure de fin prévue, sans autorisation du superviseur.", gravite: "leger", sanction: "Rappel oral puis avertissement écrit en cas de récidive", article: "RI – Art. Assiduité (à compléter)", baseLegale: "C. trav. art. L1331-1" },
];

const ECHELLE_SANCTIONS = [
  { niveau: "Recadrage oral", detail: "Rappel informel des règles, non tracé au dossier disciplinaire. Étape pédagogique avant toute sanction." },
  { niveau: "Avertissement écrit", detail: "Sanction disciplinaire la plus légère, notifiée par écrit. Pas d'entretien préalable obligatoire en principe, mais recommandé si le règlement intérieur l'impose." },
  { niveau: "Mise à pied disciplinaire", detail: "Suspension temporaire du contrat sans rémunération (durée à fixer selon le règlement intérieur). Entretien préalable obligatoire." },
  { niveau: "Rétrogradation", detail: "Modification du contrat de travail nécessitant l'accord du salarié. Entretien préalable obligatoire." },
  { niveau: "Licenciement pour cause réelle et sérieuse", detail: "Rupture du contrat avec préavis et indemnités. Entretien préalable obligatoire." },
  { niveau: "Licenciement pour faute grave", detail: "Rupture immédiate, sans préavis ni indemnité de licenciement. Entretien préalable obligatoire, mise à pied conservatoire possible." },
  { niveau: "Licenciement pour faute lourde", detail: "Faute grave avec intention de nuire à l'employeur. Mêmes effets que la faute grave, procédure renforcée." },
];

const DELAIS_LEGAUX = [
  { etape: "Délai entre les faits et la convocation", detail: "2 mois maximum à compter de la connaissance des faits par l'employeur (prescription), sauf faits répétés ou poursuites pénales." },
  { etape: "Délai de convocation → entretien", detail: "5 jours ouvrables minimum entre la remise de la convocation et la date de l'entretien préalable." },
  { etape: "Délai entretien → notification", detail: "2 jours ouvrables minimum et 1 mois maximum après l'entretien préalable pour notifier la sanction (hors avertissement simple)." },
  { etape: "Mise à pied conservatoire", detail: "Peut être prononcée immédiatement en cas de faute grave présumée, dans l'attente de la procédure disciplinaire." },
];

/* ============================================================
   MODÈLES DE COURRIERS
   Modifiables dans l'onglet "Modèles" (Hyperviseur / RH).
   Les variables entre doubles accolades {{...}} sont remplacées
   automatiquement au moment de la génération du courrier.
   ============================================================ */

const DEFAULT_MODELES = [
  {
    id: "avertissement",
    label: "Avertissement écrit",
    intro: "Sanction légère, notifiée directement par écrit.",
    champs: ["dateFaits", "descriptionFaits", "dateCourrier", "signataire"],
    template: `{{villeDate}}

{{nomCollaborateur}}
{{poste}}

Objet : Avertissement

{{civilite}},

Nous revenons vers vous au sujet des faits survenus le {{dateFaits}}, à savoir :

{{descriptionFaits}}

Ce comportement constitue un manquement à vos obligations professionnelles et contrevient aux règles applicables au sein de l'entreprise ({{articleRegle}}).

Nous vous rappelons que ce type de comportement n'est pas acceptable et ne doit pas se reproduire. Nous vous informons donc, par la présente, que nous prononçons à votre encontre un AVERTISSEMENT.

Nous espérons que cet avertissement vous permettra de prendre la mesure de la situation, et restons à votre disposition pour tout échange à ce sujet.

{{signataire}}`,
  },
  {
    id: "convocation",
    label: "Convocation à entretien préalable",
    intro: "À adresser avant toute sanction lourde (mise à pied, rétrogradation, licenciement).",
    champs: ["dateFaits", "descriptionFaits", "dateEntretien", "heureEntretien", "lieuEntretien", "miseAPiedConservatoire", "dateCourrier", "signataire"],
    template: `{{villeDate}}

{{nomCollaborateur}}
{{poste}}

Objet : Convocation à un entretien préalable à sanction{{objetMiseAPied}}
{{lrarLigne}}

{{civilite}},

Nous sommes conduits à envisager à votre encontre une sanction pouvant aller jusqu'au licenciement, en raison des faits suivants, survenus le {{dateFaits}} :

{{descriptionFaits}}

Conformément aux dispositions de l'article L.1332-2 du Code du travail, nous vous convoquons à un entretien préalable qui se tiendra :

Le {{dateEntretien}} à {{heureEntretien}}
{{lieuEntretien}}

Au cours de cet entretien, vous serez entendu(e) sur les faits qui vous sont reprochés et pourrez apporter toutes les explications utiles. Vous avez la possibilité de vous faire assister par une personne de votre choix appartenant au personnel de l'entreprise, ou, en l'absence d'institution représentative du personnel, par un conseiller extérieur dont la liste est disponible auprès de l'inspection du travail ou de la mairie.

{{miseAPiedConservatoireTexte}}

{{signataire}}`,
  },
  {
    id: "mise_a_pied",
    label: "Notification de mise à pied disciplinaire",
    intro: "À envoyer après l'entretien préalable, en respectant le délai légal.",
    champs: ["dateFaits", "descriptionFaits", "dateEntretien", "dureeMiseAPied", "dateDebutMAP", "dateFinMAP", "dateCourrier", "signataire"],
    template: `{{villeDate}}

{{nomCollaborateur}}
{{poste}}

Objet : Notification de mise à pied disciplinaire

{{civilite}},

Nous faisons suite à l'entretien préalable qui s'est tenu le {{dateEntretien}}, au cours duquel nous vous avons exposé les faits qui vous sont reprochés, survenus le {{dateFaits}} :

{{descriptionFaits}}

Les explications recueillies lors de cet entretien ne nous ont pas permis de modifier notre appréciation des faits. En conséquence, nous vous notifions par la présente une mise à pied disciplinaire de {{dureeMiseAPied}} jour(s), qui s'exécutera du {{dateDebutMAP}} au {{dateFinMAP}} inclus.

Pendant cette période, votre contrat de travail est suspendu et vous ne percevrez pas de rémunération.

Nous espérons que cette sanction vous permettra de prendre la mesure des faits reprochés et souhaitons que ceux-ci ne se reproduisent pas.

{{signataire}}`,
  },
  {
    id: "licenciement_reelle_serieuse",
    label: "Licenciement pour cause réelle et sérieuse",
    intro: "Rupture du contrat avec préavis, après entretien préalable.",
    champs: ["dateFaits", "descriptionFaits", "dateEntretien", "dureePreavis", "dateCourrier", "signataire"],
    template: `{{villeDate}}

{{nomCollaborateur}}
{{poste}}
{{lrarLigne}}

Objet : Notification de licenciement pour cause réelle et sérieuse

{{civilite}},

Nous faisons suite à l'entretien préalable qui s'est tenu le {{dateEntretien}}, au cours duquel nous vous avons exposé les motifs pour lesquels nous envisagions une mesure de licenciement, à savoir les faits survenus le {{dateFaits}} :

{{descriptionFaits}}

Les explications que vous nous avez apportées ne nous ont pas permis de revenir sur notre décision. Nous vous notifions donc, par la présente, votre licenciement pour cause réelle et sérieuse.

Votre contrat de travail prendra fin à l'issue d'un préavis de {{dureePreavis}}, débutant à la première présentation de cette lettre, sauf dispense de notre part.

Vous recevrez, à l'expiration de votre contrat, votre certificat de travail, votre reçu pour solde de tout compte, ainsi que l'attestation destinée à France Travail.

{{signataire}}`,
  },
  {
    id: "licenciement_faute_grave",
    label: "Licenciement pour faute grave",
    intro: "Rupture immédiate, sans préavis ni indemnité de licenciement.",
    champs: ["dateFaits", "descriptionFaits", "dateEntretien", "dateCourrier", "signataire"],
    template: `{{villeDate}}

{{nomCollaborateur}}
{{poste}}
{{lrarLigne}}

Objet : Notification de licenciement pour faute grave

{{civilite}},

Nous faisons suite à l'entretien préalable qui s'est tenu le {{dateEntretien}}, au cours duquel nous vous avons exposé les faits qui vous sont reprochés, survenus le {{dateFaits}} :

{{descriptionFaits}}

Ces faits constituent une violation des obligations résultant de votre contrat de travail d'une importance telle qu'ils rendent impossible votre maintien dans l'entreprise, y compris pendant la durée du préavis. Ils caractérisent une faute grave.

En conséquence, nous vous notifions par la présente votre licenciement pour faute grave, sans indemnité de licenciement ni de préavis. Votre contrat de travail prendra donc fin à la date de première présentation de cette lettre{{miseAPiedConfirmeeTexte}}.

Vous recevrez, à l'expiration de votre contrat, votre certificat de travail, votre reçu pour solde de tout compte, ainsi que l'attestation destinée à France Travail.

{{signataire}}`,
  },
];

function renderTemplate(str, data) {
  return String(str || "").replace(/\{\{(\w+)\}\}/g, (_, key) =>
    data[key] !== undefined && data[key] !== null && data[key] !== "" ? data[key] : ""
  );
}


const CHAMP_LABELS = {
  dateFaits: "Date des faits",
  descriptionFaits: "Description des faits",
  dateEntretien: "Date de l'entretien préalable",
  heureEntretien: "Heure de l'entretien",
  lieuEntretien: "Lieu de l'entretien",
  miseAPiedConservatoire: "Mise à pied conservatoire ?",
  dureeMiseAPied: "Durée de la mise à pied (jours)",
  dateDebutMAP: "Début de la mise à pied",
  dateFinMAP: "Fin de la mise à pied",
  dureePreavis: "Durée du préavis",
  dateCourrier: "Date du courrier",
  signataire: "Signataire",
};

/* ============================================================
   UTILITAIRES
   ============================================================ */

function uid(prefix) {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`;
}

function formatDateFR(iso) {
  if (!iso) return "";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString("fr-FR", { day: "numeric", month: "long", year: "numeric" });
}

function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

function normalizeDateISO(v) {
  if (v === "" || v === null || v === undefined) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number") return excelSerialToDate(v).toISOString().slice(0, 10);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  return s;
}

function normalizeTimeMinutes(v) {
  if (v === "" || v === null || v === undefined) return null;
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  if (typeof v === "number") {
    if (v >= 0 && v < 1) return Math.round(v * 1440);
    if (v >= 1) return Math.round((v - Math.floor(v)) * 1440);
    return null;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[:h](\d{2})/i);
  if (m) return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  return null;
}

function minutesToHHMM(mins) {
  if (mins === null || mins === undefined || isNaN(mins)) return "";
  const h = Math.floor(mins / 60).toString().padStart(2, "0");
  const m = (mins % 60).toString().padStart(2, "0");
  return `${h}:${m}`;
}

function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function makeCas(collaborateurId, manquementId, dateFaits, description, source) {
  return {
    id: uid("cas"),
    collaborateurId,
    manquementId,
    dateFaits,
    description,
    sanction: "",
    statut: "ouvert",
    dateCreation: new Date().toISOString(),
    source: source || "manuel",
  };
}

const GraviteBadge = ({ g }) => {
  const cls = {
    leger: "badge-leger",
    serieux: "badge-serieux",
    grave: "badge-grave",
    lourd: "badge-lourd",
  }[g];
  return <span className={`stamp ${cls}`}>{GRAVITE[g].label}</span>;
};

/* ============================================================
   ACCÈS PAR UTILISATEUR (filtre léger — voir note à l'écran de connexion)
   ============================================================ */

const ROLES = [
  { id: "superviseur", label: "Superviseur" },
  { id: "hyperviseur", label: "Hyperviseur" },
  { id: "rh", label: "RH" },
];

const DEFAULT_UTILISATEURS = [
  { id: "usr_admin_rh", nom: "Administrateur RH", role: "rh", pin: "0000" },
];

const SEUILS_DEFAUT = { retard: 5, pause: 5 };

const TAB_META = [
  { id: "aujourdhui", code: "JOUR", label: "Aujourd'hui", icon: CalendarDays, roles: ["superviseur", "hyperviseur", "rh"] },
  { id: "repertoire", code: "REP", label: "Répertoire", icon: BookOpen, roles: ["superviseur", "hyperviseur", "rh"] },
  { id: "collaborateurs", code: "COL", label: "Collaborateurs", icon: Users, roles: ["superviseur", "hyperviseur", "rh"] },
  { id: "courriers", code: "COU", label: "Courriers", icon: FileText, roles: ["superviseur", "hyperviseur", "rh"] },
  { id: "modeles", code: "MOD", label: "Modèles", icon: Settings, roles: ["hyperviseur", "rh"] },
  { id: "utilisateurs", code: "USR", label: "Utilisateurs", icon: Lock, roles: ["rh"] },
];

function LoginGate({ utilisateurs, onLogin }) {
  const [userId, setUserId] = useState(utilisateurs[0]?.id || "");
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");

  function submit(e) {
    e.preventDefault();
    const u = utilisateurs.find((x) => x.id === userId);
    if (!u) { setError("Utilisateur introuvable."); return; }
    if (String(u.pin) !== String(pin)) { setError("Code incorrect."); return; }
    setError("");
    onLogin(u);
  }

  return (
    <div className="login-screen">
      <div className="login-card">
        <div className="crest login-crest">RD</div>
        <h1>Registre disciplinaire</h1>
        <p className="subtitle">3 Media · Équipe EDF</p>
        <form onSubmit={submit} className="login-form">
          <label>
            Utilisateur
            <select value={userId} onChange={(e) => setUserId(e.target.value)}>
              {utilisateurs.map((u) => (
                <option key={u.id} value={u.id}>{u.nom} — {ROLES.find((r) => r.id === u.role)?.label}</option>
              ))}
            </select>
          </label>
          <label>
            Code d'accès
            <input type="password" inputMode="numeric" value={pin} onChange={(e) => setPin(e.target.value)} placeholder="••••" autoFocus />
          </label>
          {error && <p className="login-error">{error}</p>}
          <button type="submit" className="btn-primary-small login-submit"><Lock size={14} /> Se connecter</button>
        </form>
        <p className="login-note">
          Accès simplifié par code, propre à cet outil : il sert à séparer les usages entre superviseurs, hyperviseurs
          et RH, mais ne constitue pas une sécurité informatique robuste. Ne l'utilisez pas comme seule protection pour
          des données très sensibles — rapprochez-vous de votre service IT pour une solution avec authentification réelle.
        </p>
      </div>
    </div>
  );
}

/* ============================================================
   APP
   ============================================================ */

export default function App() {
  const [tab, setTab] = useState("aujourdhui");
  const [collaborateurs, setCollaborateurs] = useState([]);
  const [cas, setCas] = useState([]);
  const [seuils, setSeuils] = useState(SEUILS_DEFAUT);
  const [modeles, setModeles] = useState(DEFAULT_MODELES);
  const [utilisateurs, setUtilisateurs] = useState(DEFAULT_UTILISATEURS);
  const [currentUser, setCurrentUser] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved | error
  const saveTimer = useRef(null);

  // ---------- chargement ----------
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("registre-data");
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          setCollaborateurs(parsed.collaborateurs || []);
          setCas(parsed.cas || []);
          setSeuils(parsed.seuils || SEUILS_DEFAUT);
          setModeles(parsed.modeles && parsed.modeles.length ? parsed.modeles : DEFAULT_MODELES);
          setUtilisateurs(parsed.utilisateurs && parsed.utilisateurs.length ? parsed.utilisateurs : DEFAULT_UTILISATEURS);
        }
      } catch (e) {
        // pas de données existantes, on démarre avec les valeurs par défaut
      } finally {
        setLoaded(true);
      }
    })();
    try {
      const savedUserId = sessionStorage.getItem("registre-session-user");
      if (savedUserId) {
        // sera résolu une fois les utilisateurs chargés (cf. effet ci-dessous)
      }
    } catch (e) {}
  }, []);

  // ---------- restauration de session (le temps que les utilisateurs soient chargés) ----------
  useEffect(() => {
    if (!loaded || currentUser) return;
    try {
      const savedUserId = sessionStorage.getItem("registre-session-user");
      if (savedUserId) {
        const u = utilisateurs.find((x) => x.id === savedUserId);
        if (u) setCurrentUser(u);
      }
    } catch (e) {}
  }, [loaded, utilisateurs]); // eslint-disable-line react-hooks/exhaustive-deps

  // ---------- sauvegarde (debounce) ----------
  useEffect(() => {
    if (!loaded) return;
    setSaveState("saving");
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        const res = await window.storage.set(
          "registre-data",
          JSON.stringify({ collaborateurs, cas, seuils, modeles, utilisateurs }),
          false
        );
        setSaveState(res ? "saved" : "error");
      } catch (e) {
        setSaveState("error");
      }
    }, 500);
    return () => clearTimeout(saveTimer.current);
  }, [collaborateurs, cas, seuils, modeles, utilisateurs, loaded]);

  function handleLogin(u) {
    setCurrentUser(u);
    try { sessionStorage.setItem("registre-session-user", u.id); } catch (e) {}
  }
  function handleLogout() {
    setCurrentUser(null);
    try { sessionStorage.removeItem("registre-session-user"); } catch (e) {}
  }

  const visibleTabs = TAB_META.filter((t) => currentUser && t.roles.includes(currentUser.role));

  useEffect(() => {
    if (currentUser && !visibleTabs.some((t) => t.id === tab)) {
      setTab("aujourdhui");
    }
  }, [currentUser]); // eslint-disable-line react-hooks/exhaustive-deps

  if (!loaded) {
    return (
      <div className="app-scope">
        <style>{CSS}</style>
        <p style={{ padding: 40, fontFamily: "sans-serif", color: "#5C5F84" }}>Chargement…</p>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div className="app-scope">
        <style>{CSS}</style>
        <LoginGate utilisateurs={utilisateurs} onLogin={handleLogin} />
      </div>
    );
  }

  return (
    <div className="app-scope">
      <style>{CSS}</style>

      <div className="shell">
        <header className="masthead">
          <div className="masthead-left">
            <div className="crest">RD</div>
            <div>
              <h1>Registre disciplinaire</h1>
              <p className="subtitle">3 Media · Équipe EDF — répertoire des manquements &amp; suivi des procédures</p>
            </div>
          </div>
          <div className="header-right">
            <div className={`save-indicator save-${saveState}`}>
              {saveState === "saving" && "Enregistrement…"}
              {saveState === "saved" && "Enregistré"}
              {saveState === "error" && "Erreur d'enregistrement"}
            </div>
            <div className="user-chip">
              <span>{currentUser.nom}</span>
              <span className="user-chip-role">{ROLES.find((r) => r.id === currentUser.role)?.label}</span>
            </div>
            <button className="btn-quiet" onClick={handleLogout}><LogOut size={14} /> Déconnexion</button>
          </div>
        </header>

        <div className="body">
          <nav className="tabs">
            {visibleTabs.map((t) => (
              <TabButton key={t.id} id={t.id} code={t.code} icon={t.icon} active={tab === t.id} onClick={setTab}>
                {t.label}
              </TabButton>
            ))}
          </nav>

          <main className="panel">
            {tab === "aujourdhui" && (
              <Aujourdhui
                collaborateurs={collaborateurs}
                setCollaborateurs={setCollaborateurs}
                cas={cas}
                setCas={setCas}
                seuils={seuils}
                setSeuils={setSeuils}
              />
            )}
            {tab === "repertoire" && <Repertoire />}
            {tab === "collaborateurs" && (
              <Collaborateurs
                collaborateurs={collaborateurs}
                setCollaborateurs={setCollaborateurs}
                cas={cas}
                setCas={setCas}
              />
            )}
            {tab === "courriers" && (
              <Courriers collaborateurs={collaborateurs} cas={cas} setCas={setCas} modeles={modeles} currentUser={currentUser} />
            )}
            {tab === "modeles" && <GestionModeles modeles={modeles} setModeles={setModeles} />}
            {tab === "utilisateurs" && (
              <GestionUtilisateurs utilisateurs={utilisateurs} setUtilisateurs={setUtilisateurs} currentUser={currentUser} />
            )}
          </main>
        </div>

        <footer className="foot-note">
          Contenu de base générique, non contractuel — à faire valider par votre service juridique / RH et à
          personnaliser avec le règlement intérieur 3 Media avant tout usage réel.
        </footer>
      </div>
    </div>
  );
}

function TabButton({ id, code, icon: Icon, active, onClick, children }) {
  return (
    <button className={`tab-btn ${active ? "tab-btn-active" : ""}`} onClick={() => onClick(id)}>
      <span className="tab-code">{code}</span>
      <Icon size={16} strokeWidth={2} />
      <span>{children}</span>
    </button>
  );
}

/* ============================================================
   GESTION DES UTILISATEURS (RH)
   ============================================================ */

function GestionUtilisateurs({ utilisateurs, setUtilisateurs, currentUser }) {
  const [form, setForm] = useState({ nom: "", role: "superviseur", pin: "" });

  function addUser() {
    if (!form.nom.trim() || !/^\d{4,6}$/.test(form.pin)) return;
    setUtilisateurs((prev) => [...prev, { id: uid("usr"), nom: form.nom.trim(), role: form.role, pin: form.pin }]);
    setForm({ nom: "", role: "superviseur", pin: "" });
  }

  function removeUser(id) {
    if (utilisateurs.length <= 1) return;
    setUtilisateurs((prev) => prev.filter((u) => u.id !== id));
  }

  function updatePin(id, pin) {
    setUtilisateurs((prev) => prev.map((u) => (u.id === id ? { ...u, pin } : u)));
  }
  function updateRole(id, role) {
    setUtilisateurs((prev) => prev.map((u) => (u.id === id ? { ...u, role } : u)));
  }

  return (
    <div>
      <div className="panel-head">
        <div>
          <p className="eyebrow">RÉF. USR-01</p>
          <h2>Gestion des utilisateurs</h2>
        </div>
      </div>
      <p className="hint span-2" style={{ marginBottom: 16, display: "block" }}>
        Réservé au rôle RH. Rappel : ce contrôle d'accès est un filtre léger d'usage, pas une sécurité informatique
        robuste — voir la note affichée à l'écran de connexion.
      </p>

      <div className="mini-form" style={{ maxWidth: 420 }}>
        <input placeholder="Nom et prénom" value={form.nom} onChange={(e) => setForm({ ...form, nom: e.target.value })} />
        <select value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>
          {ROLES.map((r) => (<option key={r.id} value={r.id}>{r.label}</option>))}
        </select>
        <input placeholder="Code d'accès (4 à 6 chiffres)" value={form.pin} onChange={(e) => setForm({ ...form, pin: e.target.value.replace(/\D/g, "") })} maxLength={6} />
        <button className="btn-primary-small" onClick={addUser}><Plus size={14} /> Ajouter l'utilisateur</button>
      </div>

      <div className="cas-list" style={{ marginTop: 20 }}>
        {utilisateurs.map((u) => (
          <div className="cas-card" key={u.id}>
            <div className="cas-card-head">
              <div>
                <span className="manquement-libelle">{u.nom}</span>
                {u.id === currentUser.id && <span className="stamp badge-leger">Vous</span>}
              </div>
              <button className="icon-btn" onClick={() => removeUser(u.id)} disabled={utilisateurs.length <= 1}><Trash2 size={14} /></button>
            </div>
            <div className="user-edit-row">
              <label>
                Rôle
                <select value={u.role} onChange={(e) => updateRole(u.id, e.target.value)}>
                  {ROLES.map((r) => (<option key={r.id} value={r.id}>{r.label}</option>))}
                </select>
              </label>
              <label>
                Code d'accès
                <input value={u.pin} onChange={(e) => updatePin(u.id, e.target.value.replace(/\D/g, ""))} maxLength={6} />
              </label>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================================================
   GESTION DES MODÈLES DE COURRIERS (Hyperviseur / RH)
   ============================================================ */

function GestionModeles({ modeles, setModeles }) {
  const [openId, setOpenId] = useState(null);

  function updateModele(id, patch) {
    setModeles((prev) => prev.map((m) => (m.id === id ? { ...m, ...patch } : m)));
  }
  function removeModele(id) {
    setModeles((prev) => prev.filter((m) => m.id !== id));
  }
  function addModele() {
    const nm = {
      id: uid("mod"),
      label: "Nouveau modèle",
      intro: "",
      champs: ["dateFaits", "descriptionFaits", "dateCourrier", "signataire"],
      template: "{{villeDate}}\n\n{{nomCollaborateur}}\n{{poste}}\n\n{{civilite}},\n\n[Rédigez votre courrier ici — variables disponibles ci-dessous]\n\n{{signataire}}",
    };
    setModeles((prev) => [...prev, nm]);
    setOpenId(nm.id);
  }

  return (
    <div>
      <div className="panel-head">
        <div>
          <p className="eyebrow">RÉF. MOD-01</p>
          <h2>Modèles de courriers</h2>
        </div>
        <button className="btn-quiet" onClick={addModele}><Plus size={15} /> Nouveau modèle</button>
      </div>
      <p className="hint" style={{ marginBottom: 16, display: "block" }}>
        Variables disponibles : {"{{villeDate}}"}, {"{{nomCollaborateur}}"}, {"{{poste}}"}, {"{{civilite}}"}, {"{{dateFaits}}"},{" "}
        {"{{descriptionFaits}}"}, {"{{dateEntretien}}"}, {"{{heureEntretien}}"}, {"{{lieuEntretien}}"}, {"{{dureeMiseAPied}}"},{" "}
        {"{{dateDebutMAP}}"}, {"{{dateFinMAP}}"}, {"{{dureePreavis}}"}, {"{{signataire}}"}.
      </p>

      {modeles.map((m) => (
        <div className="manquement-card" key={m.id}>
          <button className="manquement-head" onClick={() => setOpenId(openId === m.id ? null : m.id)}>
            {openId === m.id ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
            <span className="manquement-libelle">{m.label}</span>
          </button>
          {openId === m.id && (
            <div className="manquement-body" style={{ paddingLeft: 16 }}>
              <label className="field-block">
                Nom du modèle
                <input value={m.label} onChange={(e) => updateModele(m.id, { label: e.target.value })} />
              </label>
              <label className="field-block">
                Description courte
                <input value={m.intro} onChange={(e) => updateModele(m.id, { intro: e.target.value })} />
              </label>
              <label className="field-block">
                Contenu du courrier
                <textarea rows={14} className="mono-textarea" value={m.template} onChange={(e) => updateModele(m.id, { template: e.target.value })} />
              </label>
              <button className="btn-danger-quiet" onClick={() => removeModele(m.id)}><Trash2 size={14} /> Supprimer ce modèle</button>
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

/* ============================================================
   VUE DU JOUR (import de présences + saisie manuelle)
   ============================================================ */

function Aujourdhui({ collaborateurs, setCollaborateurs, cas, setCas, seuils, setSeuils }) {
  const [selectedDate, setSelectedDate] = useState(todayISO());
  const [showSeuils, setShowSeuils] = useState(false);
  const [importMsg, setImportMsg] = useState("");
  const [quick, setQuick] = useState({ collaborateurId: "", nouveauNom: "", nouveauPrenom: "", manquementId: "m20", description: "" });
  const fileInputRef = useRef(null);

  const casJour = useMemo(() => {
    return cas
      .filter((c) => c.dateFaits === selectedDate)
      .map((c) => ({
        ...c,
        collaborateur: collaborateurs.find((x) => x.id === c.collaborateurId),
        manquement: MANQUEMENTS.find((m) => m.id === c.manquementId),
      }))
      .sort((a, b) => (a.collaborateur?.nom || "").localeCompare(b.collaborateur?.nom || ""));
  }, [cas, collaborateurs, selectedDate]);

  const nbConcernes = new Set(casJour.map((c) => c.collaborateurId)).size;
  const nbAttention = casJour.filter((c) => c.manquement && (c.manquement.gravite === "grave" || c.manquement.gravite === "lourd")).length;

  function addQuickCas() {
    let collabId = quick.collaborateurId;
    if (!collabId) {
      if (!quick.nouveauNom.trim() && !quick.nouveauPrenom.trim()) return;
      const nc = { id: uid("col"), nom: quick.nouveauNom.trim() || "?", prenom: quick.nouveauPrenom.trim() || "?", poste: "" };
      setCollaborateurs((prev) => [...prev, nc]);
      collabId = nc.id;
    }
    const m = MANQUEMENTS.find((x) => x.id === quick.manquementId);
    const nc = makeCas(collabId, quick.manquementId, selectedDate, quick.description || (m ? m.description : ""), "manuel");
    setCas((prev) => [...prev, nc]);
    setQuick({ collaborateurId: "", nouveauNom: "", nouveauPrenom: "", manquementId: "m20", description: "" });
  }

  function removeCas(id) {
    setCas((prev) => prev.filter((c) => c.id !== id));
  }

  async function handleFile(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setImportMsg("Import en cours…");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!rows.length) { setImportMsg("Le fichier semble vide."); return; }

      const norm = (s) => s.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
      const headerKeys = Object.keys(rows[0]);
      const colMap = {};
      headerKeys.forEach((h) => {
        const n = norm(h);
        if (n.includes("prenom")) colMap.prenom = h;
        else if (n.includes("nom")) colMap.nom = h;
        else if (n.includes("date")) colMap.date = h;
        else if (n.includes("arriv") && (n.includes("prevu") || n.includes("theoriq"))) colMap.arriveePrevue = h;
        else if (n.includes("arriv") && n.includes("reel")) colMap.arriveeReelle = h;
        else if (n.includes("depart") && (n.includes("prevu") || n.includes("theoriq"))) colMap.departPrevue = h;
        else if (n.includes("depart") && n.includes("reel")) colMap.departReelle = h;
        else if (n.includes("pause") && n.includes("autoris")) colMap.pauseAutorisee = h;
        else if (n.includes("pause") && n.includes("reel")) colMap.pauseReelle = h;
      });

      if (!colMap.nom && !colMap.prenom) {
        setImportMsg("Colonnes non reconnues : le fichier doit contenir au moins une colonne Nom et/ou Prénom.");
        return;
      }

      let collabs = [...collaborateurs];
      const nouveauxCas = [];
      let nbLignes = 0, nbAnomalies = 0;

      for (const row of rows) {
        const nom = colMap.nom ? String(row[colMap.nom]).trim() : "";
        const prenom = colMap.prenom ? String(row[colMap.prenom]).trim() : "";
        if (!nom && !prenom) continue;
        nbLignes++;
        const dateISO = colMap.date ? normalizeDateISO(row[colMap.date]) : selectedDate;

        let collab = collabs.find((c) => c.nom.toLowerCase() === nom.toLowerCase() && c.prenom.toLowerCase() === prenom.toLowerCase());
        if (!collab) {
          collab = { id: uid("col"), nom: nom || "?", prenom: prenom || "?", poste: "" };
          collabs.push(collab);
        }

        const arrPrevue = colMap.arriveePrevue ? normalizeTimeMinutes(row[colMap.arriveePrevue]) : null;
        const arrReelle = colMap.arriveeReelle ? normalizeTimeMinutes(row[colMap.arriveeReelle]) : null;
        const depPrevue = colMap.departPrevue ? normalizeTimeMinutes(row[colMap.departPrevue]) : null;
        const depReelle = colMap.departReelle ? normalizeTimeMinutes(row[colMap.departReelle]) : null;
        const pauseAutorisee = colMap.pauseAutorisee ? parseFloat(row[colMap.pauseAutorisee]) : null;
        const pauseReelle = colMap.pauseReelle ? parseFloat(row[colMap.pauseReelle]) : null;

        if (arrPrevue != null && arrReelle == null) {
          nouveauxCas.push(makeCas(collab.id, "m02", dateISO, "Absence détectée à l'import (aucun pointage d'arrivée relevé).", "import"));
          nbAnomalies++;
          continue;
        }
        if (arrPrevue != null && arrReelle != null && arrReelle - arrPrevue > seuils.retard) {
          nouveauxCas.push(makeCas(collab.id, "m01", dateISO, `Retard détecté à l'import : arrivée à ${minutesToHHMM(arrReelle)} au lieu de ${minutesToHHMM(arrPrevue)} (${arrReelle - arrPrevue} min).`, "import"));
          nbAnomalies++;
        }
        if (depPrevue != null && depReelle != null && depPrevue - depReelle > seuils.retard) {
          nouveauxCas.push(makeCas(collab.id, "m21", dateISO, `Départ anticipé détecté à l'import : départ à ${minutesToHHMM(depReelle)} au lieu de ${minutesToHHMM(depPrevue)} (${depPrevue - depReelle} min).`, "import"));
          nbAnomalies++;
        }
        if (pauseAutorisee != null && pauseReelle != null && pauseReelle - pauseAutorisee > seuils.pause) {
          nouveauxCas.push(makeCas(collab.id, "m20", dateISO, `Dépassement de pause détecté à l'import : ${pauseReelle} min prises au lieu de ${pauseAutorisee} min autorisées (+${Math.round(pauseReelle - pauseAutorisee)} min).`, "import"));
          nbAnomalies++;
        }
      }

      setCollaborateurs(collabs);
      setCas((prev) => [...prev, ...nouveauxCas]);
      setImportMsg(`${nbLignes} ligne(s) traitée(s) — ${nbAnomalies} anomalie(s) détectée(s) et ajoutée(s) au registre.`);
    } catch (err) {
      setImportMsg("Impossible de lire ce fichier. Vérifiez qu'il s'agit bien d'un fichier Excel (.xlsx) ou CSV.");
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }

  return (
    <div>
      <div className="panel-head">
        <div>
          <p className="eyebrow">RÉF. JOUR-01</p>
          <h2>Vue du jour</h2>
        </div>
        <div className="head-actions">
          <input type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} />
          <button className="btn-quiet" onClick={() => setSelectedDate(todayISO())}>Aujourd'hui</button>
          <button className="btn-quiet" onClick={() => setShowSeuils((v) => !v)}><Settings size={14} /> Seuils</button>
        </div>
      </div>

      {showSeuils && (
        <div className="bareme-box">
          <label className="field-block" style={{ maxWidth: 240 }}>
            Tolérance retard / départ anticipé (min)
            <input type="number" min="0" value={seuils.retard} onChange={(e) => setSeuils({ ...seuils, retard: parseInt(e.target.value) || 0 })} />
          </label>
          <label className="field-block" style={{ maxWidth: 240 }}>
            Tolérance dépassement de pause (min)
            <input type="number" min="0" value={seuils.pause} onChange={(e) => setSeuils({ ...seuils, pause: parseInt(e.target.value) || 0 })} />
          </label>
        </div>
      )}

      <div className="daily-stats">
        <div className="stat-card"><span className="stat-num">{casJour.length}</span><span>manquement(s) le {formatDateFR(selectedDate)}</span></div>
        <div className="stat-card"><span className="stat-num">{nbConcernes}</span><span>collaborateur(s) concerné(s)</span></div>
        <div className="stat-card stat-alert"><span className="stat-num">{nbAttention}</span><span>cas grave / lourd à traiter</span></div>
      </div>

      <div className="import-box">
        <div>
          <p className="import-title">Importer un fichier de présences (Excel ou CSV)</p>
          <p className="hint">Colonnes reconnues : Nom, Prénom, Date, Heure arrivée prévue/réelle, Heure départ prévue/réelle, Pause autorisée/réelle (min).</p>
        </div>
        <label className="btn-primary-small file-btn">
          <Upload size={14} /> Choisir un fichier
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} className="file-input-hidden" />
        </label>
      </div>
      {importMsg && <p className="import-msg">{importMsg}</p>}

      <div className="cas-form" style={{ marginTop: 20 }}>
        <label>
          Collaborateur
          <select value={quick.collaborateurId} onChange={(e) => setQuick({ ...quick, collaborateurId: e.target.value })}>
            <option value="">— Nouveau collaborateur —</option>
            {collaborateurs.map((c) => (<option key={c.id} value={c.id}>{c.prenom} {c.nom}</option>))}
          </select>
        </label>
        {!quick.collaborateurId && (
          <>
            <label>
              Prénom (nouveau)
              <input value={quick.nouveauPrenom} onChange={(e) => setQuick({ ...quick, nouveauPrenom: e.target.value })} />
            </label>
            <label>
              Nom (nouveau)
              <input value={quick.nouveauNom} onChange={(e) => setQuick({ ...quick, nouveauNom: e.target.value })} />
            </label>
          </>
        )}
        <label>
          Manquement
          <select value={quick.manquementId} onChange={(e) => setQuick({ ...quick, manquementId: e.target.value })}>
            {MANQUEMENTS.map((m) => (<option key={m.id} value={m.id}>{m.libelle}</option>))}
          </select>
        </label>
        <label className="span-2">
          Détail (optionnel)
          <textarea rows={2} value={quick.description} onChange={(e) => setQuick({ ...quick, description: e.target.value })} placeholder="Ex : pause de 32 min au lieu de 20 min autorisées" />
        </label>
        <div className="mini-form-actions span-2">
          <button className="btn-primary-small" onClick={addQuickCas}><Plus size={14} /> Ajouter au registre du jour</button>
        </div>
      </div>

      <div className="cas-list" style={{ marginTop: 20 }}>
        {casJour.length === 0 && <p className="empty-state">Aucun manquement enregistré pour cette date.</p>}
        {casJour.map((c) => (
          <div className="cas-card" key={c.id}>
            <div className="cas-card-head">
              <div>
                <span className="manquement-libelle">{c.collaborateur ? `${c.collaborateur.prenom} ${c.collaborateur.nom}` : "Collaborateur supprimé"}</span>
                {c.manquement && <GraviteBadge g={c.manquement.gravite} />}
                <span className={`source-badge source-${c.source || "manuel"}`}>{c.source === "import" ? "Import" : "Manuel"}</span>
              </div>
              <button className="icon-btn" onClick={() => removeCas(c.id)}><Trash2 size={14} /></button>
            </div>
            <p className="cas-date">{c.manquement ? c.manquement.libelle : "Manquement supprimé"}</p>
            {c.description && <p className="cas-desc">{c.description}</p>}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================================================
   RÉPERTOIRE
   ============================================================ */

function Repertoire() {
  const [query, setQuery] = useState("");
  const [catFilter, setCatFilter] = useState("all");
  const [gravFilter, setGravFilter] = useState("all");
  const [openId, setOpenId] = useState(null);
  const [showBareme, setShowBareme] = useState(false);

  const filtered = useMemo(() => {
    return MANQUEMENTS.filter((m) => {
      if (catFilter !== "all" && m.categorie !== catFilter) return false;
      if (gravFilter !== "all" && m.gravite !== gravFilter) return false;
      if (query.trim()) {
        const q = query.toLowerCase();
        if (!m.libelle.toLowerCase().includes(q) && !m.description.toLowerCase().includes(q)) return false;
      }
      return true;
    });
  }, [query, catFilter, gravFilter]);

  const grouped = useMemo(() => {
    const map = {};
    for (const cat of CATEGORIES) map[cat.id] = [];
    for (const m of filtered) map[m.categorie].push(m);
    return map;
  }, [filtered]);

  return (
    <div>
      <div className="panel-head">
        <div>
          <p className="eyebrow">RÉF. REP-01</p>
          <h2>Répertoire des manquements</h2>
        </div>
        <button className="btn-quiet" onClick={() => setShowBareme((v) => !v)}>
          <Scale size={15} />
          Barème &amp; délais
          {showBareme ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
        </button>
      </div>

      {showBareme && (
        <div className="bareme-box">
          <div className="bareme-col">
            <h4>Échelle des sanctions</h4>
            <ol className="bareme-list">
              {ECHELLE_SANCTIONS.map((s) => (
                <li key={s.niveau}>
                  <strong>{s.niveau}</strong>
                  <span>{s.detail}</span>
                </li>
              ))}
            </ol>
          </div>
          <div className="bareme-col">
            <h4><Clock size={14} /> Délais légaux à respecter</h4>
            <ol className="bareme-list">
              {DELAIS_LEGAUX.map((d) => (
                <li key={d.etape}>
                  <strong>{d.etape}</strong>
                  <span>{d.detail}</span>
                </li>
              ))}
            </ol>
          </div>
        </div>
      )}

      <div className="filter-row">
        <div className="search-box">
          <Search size={15} />
          <input
            placeholder="Rechercher un manquement…"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <select value={catFilter} onChange={(e) => setCatFilter(e.target.value)}>
          <option value="all">Toutes catégories</option>
          {CATEGORIES.map((c) => (
            <option key={c.id} value={c.id}>{c.label}</option>
          ))}
        </select>
        <select value={gravFilter} onChange={(e) => setGravFilter(e.target.value)}>
          <option value="all">Toute gravité</option>
          {Object.entries(GRAVITE).map(([k, v]) => (
            <option key={k} value={k}>{v.label}</option>
          ))}
        </select>
      </div>

      <div className="repertoire-list">
        {CATEGORIES.map((cat) => {
          const items = grouped[cat.id];
          if (!items || items.length === 0) return null;
          return (
            <div className="cat-group" key={cat.id}>
              <p className="cat-label">{cat.label}</p>
              {items.map((m) => (
                <div className="manquement-card" key={m.id}>
                  <button className="manquement-head" onClick={() => setOpenId(openId === m.id ? null : m.id)}>
                    {openId === m.id ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                    <span className="manquement-libelle">{m.libelle}</span>
                    <GraviteBadge g={m.gravite} />
                  </button>
                  {openId === m.id && (
                    <div className="manquement-body">
                      <p>{m.description}</p>
                      <div className="detail-grid">
                        <div>
                          <span className="detail-label">Sanction recommandée</span>
                          <span>{m.sanction}</span>
                        </div>
                        <div>
                          <span className="detail-label">Référence règlement intérieur</span>
                          <span className="mono">{m.article}</span>
                        </div>
                        <div>
                          <span className="detail-label">Base légale indicative</span>
                          <span className="mono">{m.baseLegale}</span>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>
          );
        })}
        {filtered.length === 0 && <p className="empty-state">Aucun manquement ne correspond à cette recherche.</p>}
      </div>
    </div>
  );
}

/* ============================================================
   COLLABORATEURS
   ============================================================ */

function Collaborateurs({ collaborateurs, setCollaborateurs, cas, setCas }) {
  const [selectedId, setSelectedId] = useState(null);
  const [showAddCollab, setShowAddCollab] = useState(false);
  const [showAddCas, setShowAddCas] = useState(false);
  const [form, setForm] = useState({ nom: "", prenom: "", poste: "" });
  const [casForm, setCasForm] = useState({ manquementId: MANQUEMENTS[0].id, dateFaits: "", description: "", sanction: "", statut: "ouvert" });

  const selected = collaborateurs.find((c) => c.id === selectedId) || null;
  const casDuCollab = cas.filter((c) => c.collaborateurId === selectedId).sort((a, b) => (b.dateFaits || "").localeCompare(a.dateFaits || ""));

  function addCollaborateur() {
    if (!form.nom.trim() || !form.prenom.trim()) return;
    const nc = { id: uid("col"), nom: form.nom.trim(), prenom: form.prenom.trim(), poste: form.poste.trim() };
    setCollaborateurs((prev) => [...prev, nc]);
    setForm({ nom: "", prenom: "", poste: "" });
    setShowAddCollab(false);
    setSelectedId(nc.id);
  }

  function removeCollaborateur(id) {
    setCollaborateurs((prev) => prev.filter((c) => c.id !== id));
    setCas((prev) => prev.filter((c) => c.collaborateurId !== id));
    if (selectedId === id) setSelectedId(null);
  }

  function addCas() {
    if (!selectedId) return;
    const nc = { id: uid("cas"), collaborateurId: selectedId, ...casForm, dateCreation: new Date().toISOString() };
    setCas((prev) => [...prev, nc]);
    setCasForm({ manquementId: MANQUEMENTS[0].id, dateFaits: "", description: "", sanction: "", statut: "ouvert" });
    setShowAddCas(false);
  }

  function removeCas(id) {
    setCas((prev) => prev.filter((c) => c.id !== id));
  }

  function updateCasStatut(id, statut) {
    setCas((prev) => prev.map((c) => (c.id === id ? { ...c, statut } : c)));
  }

  return (
    <div className="collab-layout">
      <div className="collab-side">
        <div className="panel-head-small">
          <p className="eyebrow">RÉF. COL-01</p>
          <button className="btn-primary-small" onClick={() => setShowAddCollab((v) => !v)}>
            <Plus size={14} /> Ajouter
          </button>
        </div>

        {showAddCollab && (
          <div className="mini-form">
            <input placeholder="Prénom" value={form.prenom} onChange={(e) => setForm({ ...form, prenom: e.target.value })} />
            <input placeholder="Nom" value={form.nom} onChange={(e) => setForm({ ...form, nom: e.target.value })} />
            <input placeholder="Poste" value={form.poste} onChange={(e) => setForm({ ...form, poste: e.target.value })} />
            <div className="mini-form-actions">
              <button className="btn-primary-small" onClick={addCollaborateur}>Enregistrer</button>
              <button className="btn-quiet-small" onClick={() => setShowAddCollab(false)}>Annuler</button>
            </div>
          </div>
        )}

        <ul className="collab-list">
          {collaborateurs.length === 0 && <p className="empty-state">Aucun collaborateur enregistré.</p>}
          {collaborateurs.map((c) => {
            const nbCas = cas.filter((x) => x.collaborateurId === c.id).length;
            return (
              <li key={c.id} className={selectedId === c.id ? "collab-item-active" : ""}>
                <button className="collab-item" onClick={() => setSelectedId(c.id)}>
                  <span className="collab-name">{c.prenom} {c.nom}</span>
                  <span className="collab-meta">{c.poste || "—"}</span>
                  {nbCas > 0 && <span className="mini-count">{nbCas}</span>}
                </button>
              </li>
            );
          })}
        </ul>
      </div>

      <div className="collab-main">
        {!selected && <p className="empty-state">Sélectionnez ou ajoutez un collaborateur pour voir son historique.</p>}
        {selected && (
          <>
            <div className="panel-head">
              <div>
                <p className="eyebrow">DOSSIER</p>
                <h2>{selected.prenom} {selected.nom}</h2>
                <p className="subtitle">{selected.poste || "Poste non renseigné"}</p>
              </div>
              <div className="head-actions">
                <button className="btn-quiet" onClick={() => setShowAddCas((v) => !v)}>
                  <Plus size={15} /> Ajouter un cas
                </button>
                <button className="btn-danger-quiet" onClick={() => removeCollaborateur(selected.id)}>
                  <Trash2 size={15} />
                </button>
              </div>
            </div>

            {showAddCas && (
              <div className="cas-form">
                <label>
                  Manquement
                  <select value={casForm.manquementId} onChange={(e) => setCasForm({ ...casForm, manquementId: e.target.value })}>
                    {MANQUEMENTS.map((m) => (
                      <option key={m.id} value={m.id}>{m.libelle}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Date des faits
                  <input type="date" value={casForm.dateFaits} onChange={(e) => setCasForm({ ...casForm, dateFaits: e.target.value })} />
                </label>
                <label className="span-2">
                  Description des faits
                  <textarea rows={3} value={casForm.description} onChange={(e) => setCasForm({ ...casForm, description: e.target.value })} placeholder="Détail factuel et objectif des faits observés…" />
                </label>
                <label>
                  Sanction envisagée / appliquée
                  <input value={casForm.sanction} onChange={(e) => setCasForm({ ...casForm, sanction: e.target.value })} placeholder="Ex : avertissement écrit" />
                </label>
                <label>
                  Statut
                  <select value={casForm.statut} onChange={(e) => setCasForm({ ...casForm, statut: e.target.value })}>
                    <option value="ouvert">Ouvert</option>
                    <option value="en_procedure">En procédure</option>
                    <option value="clos">Clos</option>
                  </select>
                </label>
                <div className="mini-form-actions span-2">
                  <button className="btn-primary-small" onClick={addCas}>Enregistrer le cas</button>
                  <button className="btn-quiet-small" onClick={() => setShowAddCas(false)}>Annuler</button>
                </div>
              </div>
            )}

            <div className="cas-list">
              {casDuCollab.length === 0 && <p className="empty-state">Aucun cas enregistré pour ce collaborateur.</p>}
              {casDuCollab.map((c) => {
                const m = MANQUEMENTS.find((x) => x.id === c.manquementId);
                return (
                  <div className="cas-card" key={c.id}>
                    <div className="cas-card-head">
                      <div>
                        <span className="manquement-libelle">{m ? m.libelle : "Manquement supprimé"}</span>
                        {m && <GraviteBadge g={m.gravite} />}
                      </div>
                      <button className="icon-btn" onClick={() => removeCas(c.id)}><Trash2 size={14} /></button>
                    </div>
                    <p className="cas-date">{formatDateFR(c.dateFaits) || "Date non renseignée"}</p>
                    {c.description && <p className="cas-desc">{c.description}</p>}
                    {c.sanction && <p className="cas-sanction"><strong>Sanction :</strong> {c.sanction}</p>}
                    <div className="cas-statut-row">
                      <span className="detail-label">Statut</span>
                      <select value={c.statut} onChange={(e) => updateCasStatut(c.id, e.target.value)}>
                        <option value="ouvert">Ouvert</option>
                        <option value="en_procedure">En procédure</option>
                        <option value="clos">Clos</option>
                      </select>
                    </div>
                  </div>
                );
              })}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   COURRIERS
   ============================================================ */

function Courriers({ collaborateurs, cas, setCas, modeles, currentUser }) {
  const [typeId, setTypeId] = useState(modeles[0]?.id || "");
  const [collabId, setCollabId] = useState("");
  const [casId, setCasId] = useState("");
  const [fields, setFields] = useState({});
  const [copied, setCopied] = useState(false);

  const type = modeles.find((t) => t.id === typeId) || modeles[0];
  const collab = collaborateurs.find((c) => c.id === collabId);
  const casOptions = cas.filter((c) => c.collaborateurId === collabId);

  useEffect(() => {
    if (!modeles.find((m) => m.id === typeId) && modeles[0]) setTypeId(modeles[0].id);
  }, [modeles]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    setCasId("");
  }, [collabId]);

  useEffect(() => {
    // pré-remplissage à partir du cas choisi
    if (!casId) return;
    const c = cas.find((x) => x.id === casId);
    if (!c) return;
    const m = MANQUEMENTS.find((x) => x.id === c.manquementId);
    setFields((f) => ({
      ...f,
      dateFaits: c.dateFaits || f.dateFaits,
      descriptionFaits: c.description || (m ? m.description : f.descriptionFaits),
      articleRegle: m ? m.article : f.articleRegle,
    }));
  }, [casId]); // eslint-disable-line react-hooks/exhaustive-deps

  function setField(key, value) {
    setFields((f) => ({ ...f, [key]: value }));
  }

  const villeDate = `Troyes, le ${fields.dateCourrier ? formatDateFR(fields.dateCourrier) : "[DATE]"}`;
  const nomCollaborateur = collab ? `${collab.prenom} ${collab.nom}` : (fields.nomCollaborateurLibre || "[NOM DU COLLABORATEUR]");
  const civilite = fields.civilite || "Madame, Monsieur";

  const generated = useMemo(() => {
    if (!type) return "";
    const miseAPied = !!fields.miseAPiedConservatoire;
    const data = {
      ...fields,
      villeDate,
      nomCollaborateur,
      poste: collab?.poste || fields.posteLibre || "[POSTE]",
      civilite,
      dateFaits: fields.dateFaits ? formatDateFR(fields.dateFaits) : "[DATE DES FAITS]",
      descriptionFaits: fields.descriptionFaits || "[DESCRIPTION DES FAITS]",
      dateEntretien: fields.dateEntretien ? formatDateFR(fields.dateEntretien) : "[DATE ENTRETIEN]",
      heureEntretien: fields.heureEntretien || "[HEURE]",
      lieuEntretien: fields.lieuEntretien || "[LIEU DE L'ENTRETIEN]",
      dureeMiseAPied: fields.dureeMiseAPied || "[NOMBRE]",
      dateDebutMAP: fields.dateDebutMAP ? formatDateFR(fields.dateDebutMAP) : "[DATE DÉBUT]",
      dateFinMAP: fields.dateFinMAP ? formatDateFR(fields.dateFinMAP) : "[DATE FIN]",
      dureePreavis: fields.dureePreavis || "[DURÉE]",
      signataire: fields.signataire || "Le / La responsable hiérarchique",
      articleRegle: fields.articleRegle || "règlement intérieur",
      objetMiseAPied: miseAPied ? " – Mise à pied conservatoire" : "",
      lrarLigne: "Lettre recommandée avec accusé de réception",
      miseAPiedConservatoireTexte: miseAPied
        ? "Compte tenu de la gravité des faits, nous vous notifions par la présente une mise à pied à titre conservatoire, à effet immédiat et jusqu'à l'issue de la procédure."
        : "",
      miseAPiedConfirmeeTexte: miseAPied ? ", la mise à pied conservatoire notifiée précédemment se trouvant ainsi confirmée" : "",
    };
    return renderTemplate(type.template, data);
  }, [type, fields, collab, villeDate, nomCollaborateur, civilite]);

  function copyToClipboard() {
    navigator.clipboard?.writeText(generated).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    });
  }

  function handlePrint() {
    window.print();
  }

  function handleDownloadWord() {
    const escaped = generated
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body><pre style="font-family:Calibri, Arial, sans-serif; font-size:12pt; white-space:pre-wrap;">${escaped}</pre></body></html>`;
    const blob = new Blob(["\ufeff", html], { type: "application/msword" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const nomFichier = (type?.label || "courrier").toLowerCase().replace(/[^a-z0-9]+/g, "-");
    a.href = url;
    a.download = `${nomFichier}-${nomCollaborateur.toLowerCase().replace(/[^a-z0-9]+/g, "-")}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  if (!type) return <p className="empty-state">Aucun modèle de courrier disponible. Ajoutez-en un dans l'onglet Modèles.</p>;

  return (
    <div className="courrier-layout">
      <div className="courrier-form-col no-print">
        <p className="eyebrow">RÉF. COU-01</p>
        <h2>Générer un courrier</h2>

        <label className="field-block">
          Type de courrier
          <select value={typeId} onChange={(e) => setTypeId(e.target.value)}>
            {modeles.map((t) => (
              <option key={t.id} value={t.id}>{t.label}</option>
            ))}
          </select>
          <span className="hint">{type.intro}</span>
        </label>

        <label className="field-block">
          Collaborateur
          <select value={collabId} onChange={(e) => setCollabId(e.target.value)}>
            <option value="">— Sélectionner —</option>
            {collaborateurs.map((c) => (
              <option key={c.id} value={c.id}>{c.prenom} {c.nom}</option>
            ))}
          </select>
        </label>

        {collabId && casOptions.length > 0 && (
          <label className="field-block">
            Rattacher à un cas existant (optionnel)
            <select value={casId} onChange={(e) => setCasId(e.target.value)}>
              <option value="">— Aucun / saisie libre —</option>
              {casOptions.map((c) => {
                const m = MANQUEMENTS.find((x) => x.id === c.manquementId);
                return (
                  <option key={c.id} value={c.id}>
                    {formatDateFR(c.dateFaits)} — {m ? m.libelle : "Manquement"}
                  </option>
                );
              })}
            </select>
          </label>
        )}

        <label className="field-block">
          Date du courrier
          <input type="date" value={fields.dateCourrier || ""} onChange={(e) => setField("dateCourrier", e.target.value)} />
        </label>

        {type.champs.map((champ) => {
          if (champ === "miseAPiedConservatoire") {
            return (
              <label className="field-block checkbox-row" key={champ}>
                <input
                  type="checkbox"
                  checked={!!fields.miseAPiedConservatoire}
                  onChange={(e) => setField("miseAPiedConservatoire", e.target.checked)}
                />
                {CHAMP_LABELS[champ]}
              </label>
            );
          }
          if (champ === "descriptionFaits") {
            return (
              <label className="field-block" key={champ}>
                {CHAMP_LABELS[champ]}
                <textarea rows={4} value={fields[champ] || ""} onChange={(e) => setField(champ, e.target.value)} />
              </label>
            );
          }
          if (champ === "dateFaits" || champ === "dateEntretien" || champ === "dateDebutMAP" || champ === "dateFinMAP") {
            return (
              <label className="field-block" key={champ}>
                {CHAMP_LABELS[champ]}
                <input type="date" value={fields[champ] || ""} onChange={(e) => setField(champ, e.target.value)} />
              </label>
            );
          }
          return (
            <label className="field-block" key={champ}>
              {CHAMP_LABELS[champ] || champ}
              <input value={fields[champ] || ""} onChange={(e) => setField(champ, e.target.value)} />
            </label>
          );
        })}
      </div>

      <div className="courrier-preview-col">
        <div className="preview-head no-print">
          <span className="eyebrow">Aperçu</span>
          <div className="preview-actions">
            <button className="btn-quiet" onClick={handlePrint}><Printer size={14} /> Imprimer / PDF</button>
            <button className="btn-quiet" onClick={handleDownloadWord}><Download size={14} /> Word (.doc)</button>
            <button className="btn-primary-small" onClick={copyToClipboard}>
              {copied ? <Check size={14} /> : <Copy size={14} />}
              {copied ? "Copié" : "Copier"}
            </button>
          </div>
        </div>
        <textarea className="preview-text no-print" value={generated} onChange={() => {}} readOnly />
        <div id="print-area" className="print-area">{generated}</div>
        <p className="legal-warning no-print">
          <AlertTriangle size={13} /> Modèle générique — à faire relire par le service juridique / RH avant tout envoi. Vérifiez les délais légaux dans l'onglet Répertoire.
        </p>
      </div>
    </div>
  );
}

/* ============================================================
   STYLES
   ============================================================ */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

.app-scope {
  /* Palette 3Media : bleu (action principale), violet (accent secondaire), jaune (accent tertiaire) */
  --paper: #F5F4FB;
  --paper-card: #FFFFFF;
  --ink: #1C2557;
  --ink-soft: #5C5F84;
  --line: #DAD6EC;
  --line-soft: #EEECF8;
  --teal: #2A3FA0;
  --teal-dark: #1A2A78;
  --amber: #2A3FA0;
  --amber-bg: #DEE3F8;
  --brick: #7A3FA6;
  --brick-bg: #ECE0F6;
  --dark: #55205C;
  --dark-bg: #E6D3EA;
  --gold: #8E6600;
  --gold-bg: #FCEFC4;
  --alert: #A63B3B;
  --alert-bg: #F5DCDC;
  --yellow-bright: #F2B705;

  font-family: 'IBM Plex Sans', sans-serif;
  color: var(--ink);
  background: var(--paper);
  min-height: 100%;
  -webkit-font-smoothing: antialiased;
}

.app-scope * { box-sizing: border-box; }

.shell { max-width: 1180px; margin: 0 auto; padding: 28px 20px 60px; }

.masthead {
  display: flex; align-items: center; justify-content: space-between;
  gap: 16px; padding-bottom: 20px; margin-bottom: 22px;
  border-bottom: 3px solid transparent;
  border-image: linear-gradient(90deg, var(--teal) 0%, var(--brick) 55%, var(--yellow-bright) 100%) 1;
}
.masthead-left { display: flex; align-items: center; gap: 14px; }
.crest {
  width: 46px; height: 46px; border-radius: 6px;
  background: linear-gradient(135deg, var(--brick) 0%, var(--teal) 100%);
  color: var(--yellow-bright); display: flex; align-items: center; justify-content: center;
  font-family: 'Fraunces', serif; font-weight: 700; font-size: 17px; letter-spacing: 0.5px;
  flex-shrink: 0;
}
.masthead h1 {
  font-family: 'Fraunces', serif; font-weight: 600; font-size: 24px; margin: 0;
  letter-spacing: -0.2px;
}
.subtitle { color: var(--ink-soft); font-size: 13px; margin: 3px 0 0; }
.save-indicator { font-size: 12px; color: var(--ink-soft); font-family: 'IBM Plex Mono', monospace; white-space: nowrap; }
.save-error { color: var(--alert); }

.body { display: flex; gap: 20px; align-items: flex-start; }

.tabs { display: flex; flex-direction: column; gap: 6px; width: 190px; flex-shrink: 0; }
.tab-btn {
  display: flex; align-items: center; gap: 9px;
  background: transparent; border: none; cursor: pointer;
  padding: 12px 14px; font-family: 'IBM Plex Sans', sans-serif; font-size: 13.5px;
  font-weight: 500; color: var(--ink-soft); text-align: left;
  border-radius: 8px 3px 3px 8px;
  border: 1px solid transparent;
  transition: background 0.15s, color 0.15s;
}
.tab-btn:hover { background: var(--line-soft); color: var(--ink); }
.tab-btn-active {
  background: var(--paper-card); color: var(--ink);
  border: 1px solid var(--line); border-right: 3px solid var(--teal);
  box-shadow: 0 1px 2px rgba(30,42,56,0.05);
}
.tab-code {
  font-family: 'IBM Plex Mono', monospace; font-size: 10px; color: var(--teal);
  background: var(--line-soft); padding: 2px 5px; border-radius: 3px; letter-spacing: 0.5px;
}
.tab-btn-active .tab-code { background: var(--teal); color: #fff; }

.panel {
  flex: 1; min-width: 0; background: var(--paper-card);
  border: 1px solid var(--line); border-radius: 10px;
  padding: 26px 28px; box-shadow: 0 1px 3px rgba(30,42,56,0.04);
}

.eyebrow {
  font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: 1px;
  color: var(--teal); text-transform: uppercase; margin: 0 0 4px;
}
.panel-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; margin-bottom: 18px; }
.panel-head h2 { font-family: 'Fraunces', serif; font-size: 21px; font-weight: 600; margin: 0; }
.panel-head-small { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }

.btn-quiet, .btn-quiet-small {
  display: flex; align-items: center; gap: 6px; background: var(--line-soft);
  border: 1px solid var(--line); color: var(--ink); cursor: pointer;
  padding: 8px 12px; border-radius: 7px; font-size: 12.5px; font-weight: 500;
  font-family: 'IBM Plex Sans', sans-serif;
}
.btn-quiet-small { padding: 6px 10px; font-size: 12px; }
.btn-quiet:hover, .btn-quiet-small:hover { background: var(--paper); }
.btn-danger-quiet {
  display: flex; align-items: center; background: transparent; border: 1px solid var(--alert-bg);
  color: var(--alert); cursor: pointer; padding: 8px 10px; border-radius: 7px;
}
.btn-danger-quiet:hover { background: var(--alert-bg); }
.btn-primary-small {
  display: flex; align-items: center; gap: 6px; background: var(--teal); color: #fff;
  border: none; cursor: pointer; padding: 7px 13px; border-radius: 7px; font-size: 12.5px;
  font-weight: 600; font-family: 'IBM Plex Sans', sans-serif;
}
.btn-primary-small:hover { background: var(--teal-dark); }
.icon-btn { background: transparent; border: none; color: var(--ink-soft); cursor: pointer; padding: 4px; }
.icon-btn:hover { color: var(--brick); }
.head-actions { display: flex; gap: 8px; align-items: flex-start; }

/* barème */
.bareme-box { display: flex; gap: 26px; background: var(--paper); border: 1px solid var(--line); border-radius: 8px; padding: 18px 20px; margin-bottom: 18px; flex-wrap: wrap; }
.bareme-col { flex: 1; min-width: 240px; }
.bareme-col h4 { display: flex; align-items: center; gap: 6px; font-family: 'Fraunces', serif; font-size: 14.5px; margin: 0 0 10px; }
.bareme-list { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 9px; }
.bareme-list li { display: flex; flex-direction: column; font-size: 12.5px; border-left: 2px solid var(--teal); padding-left: 10px; }
.bareme-list strong { font-size: 12.5px; }
.bareme-list span { color: var(--ink-soft); }

/* filtres répertoire */
.filter-row { display: flex; gap: 10px; margin-bottom: 18px; flex-wrap: wrap; }
.search-box { display: flex; align-items: center; gap: 8px; background: var(--paper); border: 1px solid var(--line); border-radius: 7px; padding: 8px 12px; flex: 1; min-width: 200px; color: var(--ink-soft); }
.search-box input { border: none; background: transparent; outline: none; font-size: 13px; color: var(--ink); width: 100%; }
.filter-row select { border: 1px solid var(--line); background: var(--paper); border-radius: 7px; padding: 8px 10px; font-size: 12.5px; color: var(--ink); }

.repertoire-list { display: flex; flex-direction: column; gap: 20px; }
.cat-label { font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: 0.6px; text-transform: uppercase; color: var(--ink-soft); margin: 0 0 8px; }
.manquement-card { border: 1px solid var(--line-soft); border-radius: 8px; margin-bottom: 8px; overflow: hidden; }
.manquement-head { width: 100%; display: flex; align-items: center; gap: 10px; background: var(--paper); border: none; cursor: pointer; padding: 11px 14px; text-align: left; }
.manquement-libelle { font-size: 13.5px; font-weight: 500; flex: 1; }
.manquement-body { padding: 4px 14px 16px 40px; font-size: 13px; color: var(--ink-soft); }
.manquement-body p { margin: 0 0 12px; }
.detail-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
.detail-grid > div { display: flex; flex-direction: column; gap: 3px; }
.detail-label { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--teal); font-weight: 600; }
.mono { font-family: 'IBM Plex Mono', monospace; font-size: 12px; color: var(--ink); }

.stamp {
  font-family: 'IBM Plex Mono', monospace; font-size: 10px; font-weight: 600;
  text-transform: uppercase; letter-spacing: 0.6px; padding: 3px 8px;
  border-radius: 20px; flex-shrink: 0;
}
.badge-leger { background: var(--gold-bg); color: var(--gold); }
.badge-serieux { background: var(--amber-bg); color: var(--amber); }
.badge-grave { background: var(--brick-bg); color: var(--brick); }
.badge-lourd { background: var(--dark-bg); color: var(--dark); }

.empty-state { color: var(--ink-soft); font-size: 13px; padding: 20px 0; }

/* collaborateurs */
.collab-layout { display: flex; gap: 24px; }
.collab-side { width: 240px; flex-shrink: 0; border-right: 1px solid var(--line-soft); padding-right: 20px; }
.collab-main { flex: 1; min-width: 0; }
.mini-form { display: flex; flex-direction: column; gap: 8px; background: var(--paper); border: 1px solid var(--line); border-radius: 8px; padding: 12px; margin-bottom: 14px; }
.mini-form input { border: 1px solid var(--line); border-radius: 6px; padding: 7px 9px; font-size: 12.5px; background: #fff; }
.mini-form-actions { display: flex; gap: 8px; }
.collab-list { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 4px; }
.collab-item { width: 100%; display: flex; flex-direction: column; align-items: flex-start; gap: 2px; background: transparent; border: none; border-radius: 7px; padding: 9px 10px; cursor: pointer; text-align: left; position: relative; }
.collab-item:hover { background: var(--line-soft); }
.collab-item-active .collab-item { background: var(--amber-bg); }
.collab-name { font-size: 13px; font-weight: 500; }
.collab-meta { font-size: 11.5px; color: var(--ink-soft); }
.mini-count { position: absolute; right: 10px; top: 9px; font-family: 'IBM Plex Mono', monospace; font-size: 10px; background: var(--teal); color: #fff; border-radius: 20px; padding: 1px 6px; }

.cas-form { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; background: var(--paper); border: 1px solid var(--line); border-radius: 8px; padding: 16px; margin-bottom: 18px; }
.cas-form label { display: flex; flex-direction: column; gap: 5px; font-size: 12px; color: var(--ink-soft); font-weight: 500; }
.cas-form input, .cas-form select, .cas-form textarea { border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; font-size: 13px; color: var(--ink); background: #fff; font-family: inherit; }
.span-2 { grid-column: span 2; }

.cas-list { display: flex; flex-direction: column; gap: 12px; }
.cas-card { border: 1px solid var(--line-soft); border-radius: 8px; padding: 14px 16px; }
.cas-card-head { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 6px; }
.cas-card-head > div { display: flex; align-items: center; gap: 8px; }
.cas-date { font-family: 'IBM Plex Mono', monospace; font-size: 11.5px; color: var(--ink-soft); margin: 0 0 8px; }
.cas-desc { font-size: 13px; margin: 0 0 8px; color: var(--ink); }
.cas-sanction { font-size: 12.5px; margin: 0 0 10px; color: var(--ink-soft); }
.cas-statut-row { display: flex; align-items: center; gap: 8px; }
.cas-statut-row select { border: 1px solid var(--line); border-radius: 6px; padding: 5px 8px; font-size: 12px; }

/* courriers */
.courrier-layout { display: flex; gap: 26px; }
.courrier-form-col { width: 320px; flex-shrink: 0; display: flex; flex-direction: column; gap: 14px; }
.courrier-preview-col { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.field-block { display: flex; flex-direction: column; gap: 5px; font-size: 12px; font-weight: 500; color: var(--ink-soft); }
.field-block input, .field-block select, .field-block textarea { border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; font-size: 13px; color: var(--ink); background: #fff; font-family: inherit; }
.field-block .hint { font-size: 11.5px; color: var(--teal); font-weight: 400; }
.checkbox-row { flex-direction: row; align-items: center; gap: 8px; }
.checkbox-row input { width: auto; }

.preview-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
.preview-text {
  flex: 1; min-height: 480px; border: 1px solid var(--line); border-radius: 8px;
  padding: 20px 22px; font-family: 'IBM Plex Mono', monospace; font-size: 12.5px;
  line-height: 1.7; color: var(--ink); background: var(--paper); resize: vertical;
}
.legal-warning { display: flex; align-items: center; gap: 6px; font-size: 11.5px; color: var(--alert); margin: 10px 0 0; }

.foot-note { margin-top: 26px; font-size: 11.5px; color: var(--ink-soft); text-align: center; }

@media (max-width: 860px) {
  .body { flex-direction: column; }
  .tabs { flex-direction: row; width: 100%; overflow-x: auto; }
  .collab-layout, .courrier-layout { flex-direction: column; }
  .collab-side { width: 100%; border-right: none; border-bottom: 1px solid var(--line-soft); padding-right: 0; padding-bottom: 16px; }
  .courrier-form-col { width: 100%; }
  .detail-grid { grid-template-columns: 1fr; }
  .cas-form { grid-template-columns: 1fr; }
  .span-2 { grid-column: span 1; }
}

/* écran de connexion */
.login-screen { min-height: 70vh; display: flex; align-items: center; justify-content: center; padding: 20px; }
.login-card { background: var(--paper-card); border: 1px solid var(--line); border-radius: 12px; padding: 36px 34px; max-width: 380px; width: 100%; text-align: center; box-shadow: 0 4px 18px rgba(28,37,87,0.08); }
.login-crest { margin: 0 auto 16px; }
.login-card h1 { font-family: 'Fraunces', serif; font-size: 20px; margin: 0; }
.login-form { display: flex; flex-direction: column; gap: 14px; margin-top: 22px; text-align: left; }
.login-form label { display: flex; flex-direction: column; gap: 6px; font-size: 12.5px; font-weight: 500; color: var(--ink-soft); }
.login-form select, .login-form input { border: 1px solid var(--line); border-radius: 7px; padding: 10px 12px; font-size: 14px; color: var(--ink); background: #fff; }
.login-submit { justify-content: center; padding: 10px; font-size: 13.5px; margin-top: 4px; }
.login-error { color: var(--alert); font-size: 12.5px; margin: 0; }
.login-note { font-size: 11px; color: var(--ink-soft); margin: 20px 0 0; line-height: 1.5; text-align: left; }

/* en-tête utilisateur */
.header-right { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
.user-chip { display: flex; flex-direction: column; align-items: flex-end; line-height: 1.3; font-size: 12px; }
.user-chip-role { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--teal); font-weight: 600; }

/* vue du jour */
.daily-stats { display: flex; gap: 14px; margin-bottom: 18px; flex-wrap: wrap; }
.stat-card { flex: 1; min-width: 150px; background: var(--paper); border: 1px solid var(--line); border-radius: 8px; padding: 14px 16px; display: flex; flex-direction: column; gap: 4px; }
.stat-num { font-family: 'Fraunces', serif; font-size: 24px; font-weight: 600; color: var(--teal); }
.stat-card span:last-child { font-size: 11.5px; color: var(--ink-soft); }
.stat-alert .stat-num { color: var(--alert); }

.import-box { display: flex; align-items: center; justify-content: space-between; gap: 16px; background: var(--paper); border: 1px dashed var(--line); border-radius: 8px; padding: 14px 18px; flex-wrap: wrap; }
.import-title { font-weight: 600; font-size: 13px; margin: 0 0 4px; }
.import-msg { font-size: 12.5px; color: var(--teal-dark); background: var(--amber-bg); border-radius: 6px; padding: 8px 12px; margin: 10px 0 0; }
.file-btn { position: relative; overflow: hidden; }
.file-input-hidden { position: absolute; inset: 0; opacity: 0; cursor: pointer; }

.source-badge { font-family: 'IBM Plex Mono', monospace; font-size: 9.5px; text-transform: uppercase; letter-spacing: 0.5px; padding: 2px 7px; border-radius: 20px; background: var(--line-soft); color: var(--ink-soft); }
.source-import { background: var(--violet-bg, var(--brick-bg)); color: var(--brick); }

/* gestion utilisateurs / modèles */
.user-edit-row { display: flex; gap: 14px; flex-wrap: wrap; }
.user-edit-row label { display: flex; flex-direction: column; gap: 4px; font-size: 11.5px; color: var(--ink-soft); font-weight: 500; }
.user-edit-row select, .user-edit-row input { border: 1px solid var(--line); border-radius: 6px; padding: 6px 9px; font-size: 12.5px; }
.mono-textarea { font-family: 'IBM Plex Mono', monospace; font-size: 12px; line-height: 1.6; }

/* courriers : actions d'aperçu + impression */
.preview-actions { display: flex; gap: 8px; flex-wrap: wrap; }
.print-area { display: none; }

@media print {
  body * { visibility: hidden; }
  .print-area, .print-area * { visibility: visible; }
  .print-area {
    display: block; position: absolute; top: 0; left: 0; width: 100%;
    padding: 30px 40px; white-space: pre-wrap; font-family: 'IBM Plex Mono', monospace;
    font-size: 11.5pt; line-height: 1.6; color: #111;
  }
}
`;

export default App;

