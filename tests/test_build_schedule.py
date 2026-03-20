"""Integration tests for build_schedule (v0.18.0).

Tests 6-7 from the test plan (plan-eng-review 2026-03-20).
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

import pytest
from app import build_schedule, DEFAULT_WEIGHTS, DAYS


# ---------------------------------------------------------------------------
# Helpers (same as test_scheduler.py)
# ---------------------------------------------------------------------------

def student(name, grade='C', needs=None, avail=None, backup_avail=None,
            wish_teachers=None, ng_teachers=None, ng_students=None,
            ng_dates=None, fixed=None):
    return {
        'name': name,
        'grade': grade,
        'needs': needs or {'数': 1},
        'avail': avail,
        'backup_avail': backup_avail,
        'wish_teachers': wish_teachers or [],
        'ng_teachers': ng_teachers or set(),
        'ng_students': ng_students or [],
        'ng_dates': ng_dates or set(),
        'fixed': fixed or [],
    }


def week(config):
    w = {}
    for day in DAYS:
        times = ['16', '17', '18', '19', '20'] if day != '土' else ['14', '16', '17', '18']
        w[day] = {ts: config.get(day, {}).get(ts, []) for ts in times}
    return w


def find_placements(schedule, name):
    result = []
    for wi, ws in enumerate(schedule):
        for day, times in ws.items():
            for ts, booths in times.items():
                for b in booths:
                    for slot in b['slots']:
                        if slot[1] == name:
                            result.append((wi, day, ts))
    return result


def is_primary(s, day, ts):
    return s['avail'] is None or (day, ts) in s['avail']


# ---------------------------------------------------------------------------
# Test 6: 未配置数回帰テスト
# ---------------------------------------------------------------------------

class TestNoRegressionUnplaced:
    """Test 6: 新アルゴリズムで未配置数が増えていないこと (回帰テスト)"""

    def test_all_placed_when_teachers_sufficient(self):
        """
        3 teachers × 3 days, 5 students over 2 weeks.
        With plenty of capacity, all students should be placed (unplaced=[]).
        """
        wt = [
            week({
                '月': {'16': ['T1', 'T2', 'T3']},
                '火': {'17': ['T1', 'T2', 'T3']},
                '水': {'16': ['T1', 'T2', 'T3']},
            }),
            week({
                '月': {'16': ['T1', 'T2', 'T3']},
                '火': {'17': ['T1', 'T2', 'T3']},
                '水': {'16': ['T1', 'T2', 'T3']},
            }),
        ]
        skills = {
            'T1': {'中数', '中英'},
            'T2': {'中数', '中英'},
            'T3': {'中数', '中英'},
        }

        students = [
            student('S1', avail=[('月', '16'), ('火', '17')], needs={'数': 2}),
            student('S2', avail=[('火', '17'), ('水', '16')], needs={'数': 2}),
            student('S3', avail=[('月', '16'), ('水', '16')], needs={'英': 2}),
            student('S4', avail=[('月', '16'), ('火', '17'), ('水', '16')],
                    needs={'数': 1, '英': 1}),
            student('S5', avail=None, needs={'数': 2}),  # no time constraint
        ]

        schedule, unplaced, _ = build_schedule(students, wt, skills, {}, {})

        assert unplaced == [], f"Unexpected unplaced students: {unplaced}"

    def test_constrained_student_placed_before_flexible(self):
        """
        Regression: ensure constrained students are not displaced by flexible ones.

        Setup: 2 students with different numbers of preferred slots.
        S_tight has only 1 viable slot; S_flex has 4.
        Both compete for the same slot.

        Expected: S_tight gets the contested slot (placed first); S_flex
        uses an alternative slot. Neither is unplaced.
        """
        wt = [week({
            '月': {'16': ['T1']},
            '火': {'17': ['T1']},
            '水': {'16': ['T1']},
            '木': {'17': ['T1']},
        })]
        skills = {'T1': {'中数'}}

        # X fills one seat at Mon 16 (via fixed session); 1 seat remains
        x = student('X', avail=[('月', '16')], fixed=[('月', '16', '数')])

        # Tight: only wants Mon 16 (1 viable slot after X takes one seat)
        s_tight = student('S_tight', avail=[('月', '16')])

        # Flex: wants 4 slots including Mon 16 — should yield to S_tight
        s_flex = student('S_flex',
                         avail=[('月', '16'), ('火', '17'), ('水', '16'), ('木', '17')])

        # Pass s_flex before s_tight to expose any ordering bug
        schedule, unplaced, _ = build_schedule([x, s_flex, s_tight], wt, skills, {}, {})

        assert unplaced == [], \
            f"S_tight should be placed (viable_slots ordering puts it first): {unplaced}"

        tight_places = find_placements(schedule, 'S_tight')
        assert len(tight_places) == 1

        _, day_t, ts_t = tight_places[0]
        assert is_primary(s_tight, day_t, ts_t), \
            f"S_tight should be at primary slot, got ({day_t}, {ts_t})"


# ---------------------------------------------------------------------------
# Test 7: viable_slots ソート
# ---------------------------------------------------------------------------

class TestViableSlotsSort:
    """Test 7: viable_slots ソート — 制約が多い生徒が先に配置される"""

    def test_most_constrained_student_wins_contested_slot(self):
        """
        Setup
        -----
        T1 available at Mon 16, Tue 17, Wed 16, Thu 17, Fri 16.
        Student X fills 1 of 2 seats at Mon 16 (fixed session).

        Student A: avail=[Mon 16] → 1 viable slot (most constrained)
        Student B: avail=[Mon 16, Tue 17, Wed 16, Thu 17, Fri 16] → 5 viable slots

        Students passed in order [X, B, A] to expose naive ordering issues.

        Without viable_slots fix (all return 0): B processes before A
          → B takes Mon 16's last seat → A unplaced (FAIL).
        With viable_slots fix: A(1) < B(5) → A processes first
          → A gets Mon 16's last seat → B goes to alternative → both placed (PASS).

        Expected
        --------
        Both A and B placed. A at primary Mon 16.
        """
        skills = {'T1': {'中数'}}
        wt = [week({
            '月': {'16': ['T1']},
            '火': {'17': ['T1']},
            '水': {'16': ['T1']},
            '木': {'17': ['T1']},
            '金': {'16': ['T1']},
        })]

        x = student('X', fixed=[('月', '16', '数')], avail=[('月', '16')])
        a = student('A', avail=[('月', '16')])                          # 1 viable slot
        b = student('B', avail=[('月', '16'), ('火', '17'),             # 5 viable slots
                                 ('水', '16'), ('木', '17'), ('金', '16')])

        # Intentionally pass B before A to test viable_slots ordering
        schedule, unplaced, _ = build_schedule([x, b, a], wt, skills, {}, {})

        assert unplaced == [], \
            "A (most constrained) should be placed due to viable_slots ordering"

        a_places = find_placements(schedule, 'A')
        assert len(a_places) == 1

        _, day_a, ts_a = a_places[0]
        assert is_primary(a, day_a, ts_a), \
            f"A should be placed at primary Mon 16, got ({day_a}, {ts_a})"

        # B should also be placed (at an alternative primary slot)
        b_places = find_placements(schedule, 'B')
        assert len(b_places) == 1

        _, day_b, ts_b = b_places[0]
        assert is_primary(b, day_b, ts_b), \
            f"B should be placed at a primary slot, got ({day_b}, {ts_b})"
        # B should NOT be at Mon 16 (taken by A)
        assert not (day_b == '月' and ts_b == '16'), \
            "B should have yielded Mon 16 to A"
