#!/usr/bin/env python
# -*- coding: utf-8 -*-
import argparse
from collections import Counter, defaultdict

import requests
import openpyxl
from openpyxl.styles import Alignment


def get_results(id_):
    url = f"https://api.rating.chgk.net/tournaments/{id_}/results.json?includeMasksAndControversials=1"
    results = requests.get(url)
    return results.json()


def get_cat(share):
    if 0 <= share <= 0.1:
        return "a"
    if 0.1 < share <= 0.33:
        return "b"
    if 0.33 < share <= 0.66:
        return "c"
    if 0.66 < share <= 0.9:
        return "d"
    if share > 0.9:
        return "e"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("id")
    args = parser.parse_args()

    results = get_results(args.id)
    t_rating = Counter()
    t_questions = {}
    t_by_cat = defaultdict(Counter)
    teams_taken = defaultdict(list)
    n_teams = len(results)
    for res in results:
        team = (res["team"]["id"], res["current"]["name"])
        t_questions[team] = res["questionsTotal"]
        for i, v in enumerate(list(res["mask"])):
            q_num = i + 1
            if v == "1":
                teams_taken[q_num].append(team)
    all_cats = Counter()
    for q in sorted(teams_taken):
        teams = teams_taken[q]
        rating = n_teams - len(teams) + 1
        cat = get_cat(len(teams) / n_teams)
        all_cats[cat] += 1
        for t in teams:
            t_rating[t] += rating
            t_by_cat[t][cat] += 1
        print(f"Вопрос {q}: {rating}, {cat}")
    srt = sorted(t_questions, key=lambda x: (t_rating[x], t_questions[x]), reverse=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    header = [
        ("ID", 7),
        ("Команда", 25),
        ("Рейтинг", 7),
        ("Взято", 5),
        ("Ср. рейт", 7),
        ("A", 3.5),
        ("B", 3.5),
        ("C", 3.5),
        ("D", 3.5),
        ("E", 3.5),
        ("", 3.5),
        ("A", 3.5),
        ("B", 3.5),
        ("C", 3.5),
        ("D", 3.5),
        ("E", 3.5),
    ]
    for i, tup in enumerate(header):
        cell = ws.cell(row=1, column=i + 1)
        cell.value = tup[0]
        ws.column_dimensions[cell.column_letter].width = tup[1]
        if i != 1:
            cell.alignment = Alignment(horizontal="right")
    #ws.append([x[0] for x in header])
    first = True
    categs = ["a", "b", "c", "d", "e"]
    for t in srt:
        q = t_questions[t]
        r = t_rating[t]
        avg_rating = round(r / q, 2)
        row = [
            t[0],
            t[1],
            r,
            q,
            avg_rating,
        ] + [
            t_by_cat[t][c] for c in categs
        ]
        if first:
            row.extend([""] + [
                all_cats[c] for c in categs
            ])
            first = False
        ws.append(row)
    wb.save(f"{args.id}.xlsx")


if __name__ == "__main__":
    main()
