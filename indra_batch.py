# -*- coding: utf-8 -*-
"""
indra_batch.py – помощник для pdf_to_md_indra.ps1.

Получает JSON-файл с заданиями [{"src": ..., "out": ...}, ...], распознаёт все
исходники одним запуском Indra OCR (модели загружаются один раз, страницы всех
файлов распределяются по ядрам) и пишет Markdown каждого файла в указанный
"out" (временный путь – переименование в итоговый делает PowerShell).

Результат по каждому заданию записывается в JSON-файл отчёта:
  {"src", "status": "done" | "empty" | "error", "reason", "pages", "page_errors"}

Запуск:
  python indra_batch.py <jobs.json> <result.json> [--cores N] [--mode auto|ocr|text]
                        [--keep-furniture] [--dpi 200] [--chunk-pages 300]
"""

import argparse
import json
import os
import sys
import time
import traceback


def _write_result(path, results):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def _page_count(path):
    """Число страниц PDF (для разбиения на порции); для картинок – 1."""
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(path)
        n = len(doc)
        doc.close()
        return n
    except Exception:
        return 1


def _chunks(jobs, chunk_pages):
    """Порции заданий примерно по chunk_pages страниц: пул процессов
    пересоздаётся на каждую порцию, чтобы не держать в памяти результаты
    всей большой папки и не терять всё при сбое."""
    chunk, pages = [], 0
    for job in jobs:
        n = _page_count(job["src"])
        if chunk and pages + n > chunk_pages:
            yield chunk
            chunk, pages = [], 0
        chunk.append(job)
        pages += n
    if chunk:
        yield chunk


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("jobs")
    ap.add_argument("result")
    ap.add_argument("--cores", type=int, default=0)
    ap.add_argument("--mode", choices=["auto", "ocr", "text"], default="auto")
    ap.add_argument("--keep-furniture", action="store_true")
    ap.add_argument("--dpi", type=float, default=200.0)
    ap.add_argument("--chunk-pages", type=int, default=300)
    a = ap.parse_args()

    # utf-8-sig: PowerShell 5.1 пишет UTF-8 с BOM
    with open(a.jobs, "r", encoding="utf-8-sig") as f:
        jobs = json.load(f)
    if isinstance(jobs, dict):  # ConvertTo-Json свернул массив из одного элемента
        jobs = [jobs]

    import indra_ocr

    results = []
    total = len(jobs)
    k = 0
    started = time.time()

    for chunk in _chunks(jobs, a.chunk_pages):
        n_before = len(results)
        try:
            stream = indra_ocr.convert_many(
                [j["src"] for j in chunk],
                cores=a.cores or None,
                mode=a.mode,
                furniture=a.keep_furniture,
                dpi=a.dpi,
            )
            # convert_many отдаёт результаты в порядке входа
            for job, (path, result) in zip(chunk, stream):
                k += 1
                rec = {"src": job["src"], "status": "error", "reason": "",
                       "pages": 0, "page_errors": []}
                name = os.path.basename(job["src"])
                if isinstance(result, indra_ocr.InputError):
                    msg = str(result)
                    if "password" in msg.lower():
                        rec["reason"] = "PDF зашифрован или защищён паролем"
                    else:
                        rec["reason"] = "не удалось открыть файл – " + msg
                    print(f"  [{k}/{total}] ОШИБКА: {name} – {rec['reason']}", flush=True)
                    results.append(rec)
                    _write_result(a.result, results)
                    continue

                chunks_md, ocr_pages, text_pages = [], 0, 0
                for i, page in result:
                    if page.source == "error":
                        rec["page_errors"].append(f"стр. {i + 1}: {page.note}")
                    elif page.source == "text":
                        text_pages += 1
                    else:
                        ocr_pages += 1
                    chunks_md.append(page.markdown if page.markdown is not None else (page.text or ""))
                rec["pages"] = len(result)
                body = "\n\n".join(chunks_md).strip() + "\n"
                try:
                    with open(job["out"], "w", encoding="utf-8", newline="\n") as f:
                        f.write(body)
                except OSError as e:
                    rec["reason"] = f"не удалось записать результат – {e}"
                    print(f"  [{k}/{total}] ОШИБКА: {name} – {rec['reason']}", flush=True)
                    results.append(rec)
                    _write_result(a.result, results)
                    continue

                rec["status"] = "done" if body.strip() else "empty"
                if rec["page_errors"] and len(rec["page_errors"]) == rec["pages"]:
                    rec["status"] = "error"
                    rec["reason"] = "не распознана ни одна страница – " + "; ".join(rec["page_errors"][:3])
                info = f"{rec['pages']} стр.: текстовый слой {text_pages}, OCR {ocr_pages}"
                if rec["page_errors"]:
                    info += f", сбоев {len(rec['page_errors'])}"
                print(f"  [{k}/{total}] {name} ({info})", flush=True)
                results.append(rec)
                _write_result(a.result, results)
        except Exception as e:  # сбой всей порции: помечаем её необработанные файлы
            detail = f"{type(e).__name__}: {e}"
            for j in chunk[len(results) - n_before:]:  # файлы порции без результата
                k += 1
                results.append({"src": j["src"], "status": "error",
                                "reason": "сбой Indra OCR – " + detail,
                                "pages": 0, "page_errors": []})
            print("  Сбой Indra OCR на порции файлов: " + detail, flush=True)
            traceback.print_exc()
            _write_result(a.result, results)

    _write_result(a.result, results)
    print(f"  Indra OCR: {total} файл(ов) за {time.time() - started:.0f} с", flush=True)
    return 0


if __name__ == "__main__":  # обязательно: на Windows рабочие процессы импортируют этот файл
    sys.exit(main())
