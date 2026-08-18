# TGS-2024049211 — CompTIA CySA+ (CS0-003) Training

> **Course:** WSQ — CompTIA Cybersecurity Analyst (CySA+) Training
> **Course Code:** TGS-2024049211
> **Register here:** https://www.tertiarycourses.com.sg/wsq-comptia-cybersecurity-analyst-cysa-training.html

Official courseware and hands-on lab exercises for the WSQ CompTIA Cybersecurity Analyst (CySA+) Training course delivered by [**Tertiary Infotech Academy Pte Ltd**](https://www.tertiarycourses.com.sg/).

The course follows the **four CompTIA CySA+ CS0-003 exam domains** in order and weight:

| Topic | CS0-003 Exam Domain | Weight | Labs |
|---|---|---|---|
| 1 | Security Operations | 33% | 1–10 |
| 2 | Vulnerability Management | 30% | 11–19 |
| 3 | Incident Response and Management | 20% | 20–25 |
| 4 | Reporting and Communication | 17% | 26–30 |

The slide deck (v8) opens with the exam domains, carries an in-depth explanation of every CySA+ concept plus the full theory for each domain, followed by its labs, and closes with a **practice exam** and the **CompTIA CySA+ exam registration guide** (Pearson VUE).

A complete set of **30 step-by-step labs** aligned to the CompTIA CySA+ CS0-003 exam objectives. Most labs run on the free **Killercoda Ubuntu Playground** (https://killercoda.com/playgrounds/scenario/ubuntu) — no local install required. A few use free web tools or free virtual appliances.

---

## How to use

1. Open the Killercoda playground in your browser: https://killercoda.com/playgrounds/scenario/ubuntu
2. Pick a lab from the list below and follow the steps in order.
3. Reset the playground between labs that change firewall rules or install heavy services.
4. See [labs/tools.md](labs/tools.md) for every free tool used (with install commands and download links).

---

## Lab catalogue

### Domain 1 — Security Operations (33%)
- [Lab 1 — Log Ingestion and Time Synchronization](labs/lab-01-log-ingestion-time-sync/)
- [Lab 2 — OS Hardening and System Process Inspection](labs/lab-02-os-hardening/)
- [Lab 3 — Network Segmentation and Zero Trust](labs/lab-03-segmentation-zero-trust/)
- [Lab 4 — Detecting Malicious Network Activity](labs/lab-04-malicious-network-activity/)
- [Lab 5 — Host-Based IOC Hunting](labs/lab-05-host-ioc-hunting/)
- [Lab 6 — Packet Capture for Threat Hunting](labs/lab-06-packet-capture-threat-hunting/)
- [Lab 7 — SIEM Log Correlation](labs/lab-07-siem-log-correlation/)
- [Lab 8 — Email Header Analysis (SPF/DKIM/DMARC)](labs/lab-08-email-header-analysis/)
- [Lab 9 — Malware Triage with Hashing and VirusTotal](labs/lab-09-malware-triage/)
- [Lab 10 — Threat Intelligence and MITRE ATT&CK](labs/lab-10-threat-intel-mitre/)

### Domain 2 — Vulnerability Management (30%)
- [Lab 11 — Asset Discovery with Nmap](labs/lab-11-asset-discovery-nmap/)
- [Lab 12 — Vulnerability Scanning with OpenVAS](labs/lab-12-openvas-scanning/)
- [Lab 13 — Web App Scanning with OWASP ZAP](labs/lab-13-owasp-zap/)
- [Lab 14 — Web Recon with Nikto and Burp Suite](labs/lab-14-nikto-burp/)
- [Lab 15 — Metasploit Framework Basics](labs/lab-15-metasploit-basics/)
- [Lab 16 — CVSS Scoring and Prioritization](labs/lab-16-cvss-prioritization/)
- [Lab 17 — Exploiting and Mitigating XSS and SQL Injection](labs/lab-17-xss-sqli/)
- [Lab 18 — Patch Management and Hardening](labs/lab-18-patch-hardening/)
- [Lab 19 — Attack Surface Reconnaissance](labs/lab-19-attack-surface-recon/)

### Domain 3 — Incident Response and Management (20%)
- [Lab 20 — Cyber Kill Chain and ATT&CK Mapping](labs/lab-20-kill-chain-mapping/)
- [Lab 21 — Evidence Acquisition and Chain of Custody](labs/lab-21-evidence-chain-of-custody/)
- [Lab 22 — Memory Forensics with Volatility](labs/lab-22-memory-forensics/)
- [Lab 23 — Log Analysis for Incident Response](labs/lab-23-ir-log-analysis/)
- [Lab 24 — Containment with Host Isolation](labs/lab-24-containment-isolation/)
- [Lab 25 — Incident Response Playbook and Tabletop](labs/lab-25-ir-playbook-tabletop/)

### Domain 4 — Reporting and Communication (17%)
- [Lab 26 — Vulnerability Management Report](labs/lab-26-vuln-report/)
- [Lab 27 — Executive Incident Report](labs/lab-27-executive-incident-report/)
- [Lab 28 — Security Metrics Dashboard (MTTD/MTTR)](labs/lab-28-metrics-dashboard/)
- [Lab 29 — Compliance Reporting (PCI DSS / ISO 27001)](labs/lab-29-compliance-reporting/)
- [Lab 30 — Stakeholder Communication and Lessons Learned](labs/lab-30-stakeholder-lessons-learned/)

---

## Courseware

- `courseware/CompTIA Cybersecurity Analyst (CySA+) Training-v8.pptx` — Course slides (919 slides, CS0-003 domain-ordered, in-depth concept explanations)
- `courseware/LG-CompTIA Cybersecurity Analyst (CySA+) Training.docx` — Learner Guide
- `courseware/LP-CompTIA Cybersecurity Analyst (CySA+) Training.docx` — Lesson Plan (5 days)

## Reference

- [labs/README.md](labs/README.md) — Lab index grouped by domain with software requirements (each lab in its own folder)
- [labs/tools.md](labs/tools.md) — Complete list of free tools (Killercoda + external)
- `cmptia_cysaplus_cso-003.pdf` — Official exam blueprint
- [CompTIA CySA+ exam registration](https://www.comptia.org/certifications/cybersecurity-analyst)

---

## Free tools used

All tooling is **100% free**. The bulk runs inside the disposable Killercoda VM via `apt`. A few labs use free web tools (VirusTotal, AbuseIPDB, MITRE ATT&CK Navigator) or free virtual appliances (DVWA, Metasploitable, OpenVAS / Greenbone CE).

Full tool list: [labs/tools.md](labs/tools.md).
