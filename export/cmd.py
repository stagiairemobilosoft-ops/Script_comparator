from __future__ import annotations

import argparse
from datetime import datetime

from normalize.file_loader import *
from constants.constant import *
from comparison.compare import *
from constants.constant import (
    ACTION_UPDATE,
    ACTION_CLOSE,
    ACTION_CREATE,
    OUTPUT_LOG,
    ACTION_NO_CHANGE,
)
from validation.quality_check import deduplicate_stores, find_data_quality_issues
from .log import *
from .excel_export import build_sync_file, write_excel
from comparison.utils import print_stage, print_step_done

def _ensure_output_dirs() -> None:
    OUTPUT_RESULTAT.mkdir(parents=True, exist_ok=True)
    OUTPUT_LOG.mkdir(parents=True, exist_ok=True)

def cmd_compare(args: argparse.Namespace) -> int:
    show_progress = not args.quiet
    previous_path = Path(args.previous)
    incoming_path = Path(args.incoming)

    try:
        print_stage(1, 5, f"Chargement {previous_path.name}...", enabled=show_progress)
        ref_df, _ = deduplicate_stores(load_store_file(previous_path))
        print_step_done(f"{len(ref_df)} magasin(s) charges", enabled=show_progress)

        print_stage(2, 5, f"Chargement {incoming_path.name}...", enabled=show_progress)
        inc_raw = load_store_file(incoming_path)
        inc_df, inc_duplicates = deduplicate_stores(inc_raw)
        print_step_done(f"{len(inc_df)} magasin(s) charges", enabled=show_progress)

        print_stage(3, 5, "Comparaison des magasins...", enabled=show_progress)
        result = compare_store_files(ref_df, inc_df, show_progress=show_progress)
    except LoadError as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        return 1

    _ensure_output_dirs()
    output_path = Path(args.output) if args.output else DEFAULT_COMPARE_OUTPUT
    if not output_path.is_absolute() and output_path.parent == Path("."):
        output_path = OUTPUT_RESULTAT / output_path.name

    print_stage(4, 5, "Generation du fichier Filezilla...", enabled=show_progress)
    sync_file = build_sync_file(result)
    write_excel(output_path, to_result_layout(sync_file, show_progress=show_progress))
    print_step_done(f"{len(sync_file)} ligne(s) ecrites", enabled=show_progress)

    print_stage(5, 5, "Controle qualite et rapport d'analyse...", enabled=show_progress)
    quality = find_data_quality_issues(inc_raw, show_progress=show_progress)
    dup_count = inc_duplicates[STORE_ID_COLUMN].nunique() if not inc_duplicates.empty else 0

    run_stamp = timestamp()
    analysis_path = timestamped_analysis_path(run_stamp)
    write_analysis_workbook(
        analysis_path, result, previous_path, incoming_path, output_path, dup_count, inc_duplicates, quality
    )
    print_step_done("Rapport d'analyse enregistre", enabled=show_progress)
    if show_progress:
        print(file=sys.stderr)

    log_lines = [
        f"date={datetime.now().isoformat(timespec='seconds')}",
        f"commande=compare",
        f"mois_precedent={previous_path.resolve()}",
        f"nouveau={incoming_path.resolve()}",
        f"resultat={output_path.resolve()}",
        f"analyse={analysis_path.resolve()}",
        f"total_reference={result.reference_total}",
        f"total_entrant={result.incoming_total}",
        f"{ACTION_NO_CHANGE}={result.unchanged_count}",
        f"{ACTION_UPDATE}={len(result.modified)}",
        f"{ACTION_CREATE}={len(result.added)}",
        f"{ACTION_CLOSE}={len(result.removed)}",
        f"lignes_filezilla={len(sync_file)}",
        f"doublons={dup_count}",
    ]
    log_path = timestamped_log_path("compare", run_stamp)
    write_log(log_path, log_lines)

    print("Classification terminée (flow IT).")
    if dup_count:
        print(f"  Doublons entrant : {dup_count} code(s) en double (dernière ligne conservée)")
    print(f"  Mois précédent             : {result.reference_total} magasins")
    print(f"  Nouveau fichier            : {result.incoming_total} magasins")
    print(f"  {ACTION_NO_CHANGE}       : {result.unchanged_count} (exclus du fichier)")
    print(f"  {ACTION_UPDATE}          : {len(result.modified)} (status {STATUS_OPEN})")
    print(f"  {ACTION_CREATE}          : {len(result.added)} (status {STATUS_OPEN})")
    print(f"  {ACTION_CLOSE}           : {len(result.removed)} (status {STATUS_CLOSED})")
    print(f"  Lignes dans le fichier Filezilla : {len(sync_file)}")
    print(f"Résultat : {output_path.resolve()}")
    print(f"Analyse  : {analysis_path.resolve()}")
    print(f"Log      : {log_path.resolve()}")
    return 0


def cmd_validate(args: argparse.Namespace) -> int:
    try:
        df = load_store_file(Path(args.fichier))
    except LoadError as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        return 1

    _ensure_output_dirs()
    issues = find_data_quality_issues(df)
    output = Path(args.output)
    write_excel(output, issues)

    log_path = timestamped_log_path("validate")
    write_log(
        log_path,
        [
            f"date={datetime.now().isoformat(timespec='seconds')}",
            "commande=validate",
            f"fichier={Path(args.fichier).resolve()}",
            f"problemes={len(issues)}",
            f"rapport={output.resolve()}",
        ],
    )
    print(f"{len(issues)} problème(s) détecté(s). Rapport : {output.resolve()}")
    print(f"Log : {log_path.resolve()}")
    return 0
