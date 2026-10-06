# InvioReti: distribution of ERP shipment data to the sales networks by e-mail

*Invio dei dati di spedizione dal gestionale MFG alle reti di vendita*

**Visual Basic 6** · 2003 · version 7.1.0  
Author: **Massimo Sbarbaro** ([ORCID 0009-0006-8965-9013](https://orcid.org/0009-0006-8965-9013))

## Overview

InvioReti is a batch tool with a single window. It extracts the shipped orders from the MFG/PRO ERP, updates the central complaints database and sends to each sales network, by e-mail, the list of new orders shipped to its customers. These messages are the input that the program *ComplaintCard* reads in each sales office, so that complaint cards can be linked to real orders.

I designed and programmed this application in 2003. It is published here in 2026 as part of the documented record of my software work, with a permanent DOI on Zenodo.

## Main functions

- Runs an Attachmate Reflection script that queries the ERP host and produces a text export.
- Downloads the export by FTP and loads it into a staging table using a configurable field map (`CorrImportMFG`).
- Inserts or updates the `MFG` table of the central complaints database.
- Queues the new orders per sales network and sends them by SMTP; recipients are read from the `EMailxRete` table.

## Data

Microsoft Access databases through DAO (staging `$$Temp.mdb` and the central `Reclami.mdb`).

## Technology

DAO, Microsoft Access 8 object library, Winsock, Masked Edit, Common Controls.

## Repository contents

| Path | Content |
|---|---|
| `src/` | Visual Basic 6 project (`.vbp`), forms (`.frm` with their binary resources `.frx`) and modules (`.bas`), as listed in the project file. |
| `config-example/` | Templates of the `.ini` configuration files read at start-up, with placeholder values. |

## What is not included

Crystal Reports layouts (`.rpt`), compiled executables, installers, scripts for the host connection and the production databases are **not** included, because they contain operational data of the organisations for which the program was built. Configuration files are published as templates in `config-example/` with placeholder values: server addresses, accounts and passwords have been removed. Names of client organisations and products have been removed from captions and comments; local paths have been normalised.

## Related repositories

- [quality-complaints-manager-vb6](https://github.com/massimosbarbaro/quality-complaints-manager-vb6)
- [complaint-card-sales-network-vb6](https://github.com/massimosbarbaro/complaint-card-sales-network-vb6)

## How to cite

Use the citation metadata in [`CITATION.cff`](CITATION.cff) (GitHub: *Cite this repository*). Each release is archived on Zenodo with its own DOI.

> Sbarbaro, Massimo. *InvioReti: distribution of ERP shipment data to the sales networks by e-mail (Visual Basic 6, 2003)*. Software, version 7.1.0. GitHub: https://github.com/massimosbarbaro/erp-shipments-to-sales-networks-vb6

## License

Released under the [MIT License](LICENSE). © 2003 Massimo Sbarbaro.
