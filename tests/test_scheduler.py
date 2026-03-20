"""Unit tests for scheduler algorithm improvements (v0.18.0).

Tests 1-5 from the test plan (plan-eng-review 2026-03-20).
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

import pytest
from app import build_schedule, DEFAULT_WEIGHTS, DAYS


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def student(name, grade='C', needs=None, avail=None, backup_avail=None,
            wish_teachers=None, ng_teachers=None, ng_students=None,
            ng_dates=None, fixed=None):
    """Build a minimal student dict accepted by build_schedule."""
    return {
        'name': name,
        'grade': grade,
        'needs': needs or {'数': 1},
        'avail': avail,           # None = no time constraint (any slot OK)
        'backup_avail': backup_avail,
        'wish_teachers': wish_teachers or [],
        'ng_teachers': ng_teachers or set(),
        'ng_students': ng_students or [],
        'ng_dates': ng_dates or set(),
        'fixed': fixed or [],
    }


def week(config):
    """Build a weekly_teachers[wi] dict from a sparse spec.

    config: {day: {ts_short: [teacher_name, ...]}}
    Missing slots default to empty list.
    """
    w = {}
    for day in DAYS:
        times = ['16', '17', '18', '19', '20'] if day != '土' else ['14', '16', '17', '18']
        w[day] = {ts: config.get(day, {}).get(ts, []) for ts in times}
    return w


SKILLS = {'T1': {'中数', '中英'}, 'T2': {'中数', '中英'}}


def find_placements(schedule, name):
    """Return list of (week_idx, day, ts) for student across all weeks."""
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
    """True if (day, ts) is a primary (preferred) slot for student s."""
    return s['avail'] is None or (day, ts) in s['avail']


# ---------------------------------------------------------------------------
# Test 1: swap_pass() 基本動作
# ---------------------------------------------------------------------------

class TestSwapPassBasic:
    """Test 1: swap_pass() 基本動作 — backup → primary へのスワップ確認"""

    def test_backup_student_moves_to_primary_after_swap(self):
        """
        Setup
        -----
        T1 is available at Mon 16:00 and Tue 17:05.

        B (wish_teachers=[T1]) → placed first (Phase2a, wish_order)
          → takes ('月', '16') with T1 [primary for B].

        A wants ('月', '16') but has ng_students=['B']
          → cannot share T1's booth with B
          → falls back to ('火', '17') [backup for A].

        Expected after swap_pass
        ------------------------
        A at ('月', '16') primary, B at ('火', '17') primary.
        Both primary_count increases from 1 to 2.
        """
        wt = [week({'月': {'16': ['T1']}, '火': {'17': ['T1']}})]

        a = student('A',
                    avail=[('月', '16')],
                    backup_avail=[('火', '17')],
                    ng_students=['B'])
        b = student('B',
                    avail=[('月', '16'), ('火', '17')],
                    wish_teachers=['T1'])

        schedule, unplaced, _ = build_schedule([a, b], wt, SKILLS, {}, {})

        assert unplaced == [], f"Unexpected unplaced: {unplaced}"

        a_places = find_placements(schedule, 'A')
        b_places = find_placements(schedule, 'B')
        assert len(a_places) == 1
        assert len(b_places) == 1

        _, day_a, ts_a = a_places[0]
        _, day_b, ts_b = b_places[0]

        assert is_primary(a, day_a, ts_a), \
            f"A should be at primary after swap, got ({day_a}, {ts_a})"
        assert is_primary(b, day_b, ts_b), \
            f"B should still be at primary after swap, got ({day_b}, {ts_b})"


# ---------------------------------------------------------------------------
# Test 2: swap_pass() NG制約キャンセル
# ---------------------------------------------------------------------------

class TestSwapPassNGCancel:
    """Test 2: swap_pass() NG制約 — スワップが阻止されること"""

    def test_swap_blocked_by_ng_teacher(self):
        """
        Setup
        -----
        T1 only at Mon 16:00, T2 only at Tue 17:05.

        A.ng_teachers={'T1'}: A cannot use T1 → cannot be at Mon 16.
        A placed at backup ('火', '17') with T2.
        B (wish_teachers=[T1]) at primary ('月', '16') with T1.

        swap_pass tries A↔B, but check_booth for A at Mon 16 fails
        (T1 in A's ng_teachers) → swap is cancelled.

        Expected
        --------
        A remains at backup ('火', '17').
        """
        wt = [week({'月': {'16': ['T1']}, '火': {'17': ['T2']}})]
        skills = {'T1': {'中数'}, 'T2': {'中数'}}

        a = student('A',
                    avail=[('月', '16')],
                    backup_avail=[('火', '17')],
                    ng_teachers={'T1'})
        b = student('B',
                    avail=[('月', '16'), ('火', '17')],
                    wish_teachers=['T1'])

        schedule, unplaced, _ = build_schedule([a, b], wt, skills, {}, {})

        a_places = find_placements(schedule, 'A')
        assert len(a_places) == 1, f"A should be placed once, got {a_places}"

        _, day_a, ts_a = a_places[0]
        assert not is_primary(a, day_a, ts_a), \
            f"A should remain at backup (swap blocked by ng_teacher), got ({day_a}, {ts_a})"


# ---------------------------------------------------------------------------
# Test 3: swap_pass() early exit
# ---------------------------------------------------------------------------

class TestSwapPassEarlyExit:
    """Test 3: swap_pass() early exit — 全員primary時にスワップゼロで終了"""

    def test_no_swaps_when_all_primary(self):
        """
        When every placed student is already at a primary slot,
        swap_pass finds 0 beneficial swaps and exits immediately.

        Verify: all students remain at primary slots (correct result),
        and no extra placements appear.
        """
        wt = [week({'月': {'16': ['T1']}, '火': {'17': ['T1']}})]

        a = student('A', avail=[('月', '16')])
        b = student('B', avail=[('火', '17')])

        schedule, unplaced, _ = build_schedule([a, b], wt, SKILLS, {}, {})

        assert unplaced == []

        a_places = find_placements(schedule, 'A')
        b_places = find_placements(schedule, 'B')
        assert len(a_places) == 1
        assert len(b_places) == 1

        _, day_a, ts_a = a_places[0]
        _, day_b, ts_b = b_places[0]

        assert is_primary(a, day_a, ts_a), \
            f"A should stay at primary, got ({day_a}, {ts_a})"
        assert is_primary(b, day_b, ts_b), \
            f"B should stay at primary, got ({day_b}, {ts_b})"


# ---------------------------------------------------------------------------
# Test 4: swap_pass() 同週内のみ
# ---------------------------------------------------------------------------

class TestSwapPassSameWeek:
    """Test 4: swap_pass() 同週内のみ — 週をまたぐスワップは行わない"""

    def test_multi_week_students_resolved_independently(self):
        """
        2-week schedule. A needs 2 sessions (1 per week).
        swap_pass operates per-week; cross-week slots are never compared.

        Verify: A ends up at primary in BOTH weeks (swap worked independently
        in each week where needed), and no unplaced entries exist.
        """
        wt = [
            week({'月': {'16': ['T1']}, '火': {'17': ['T1']}}),
            week({'月': {'16': ['T1']}, '火': {'17': ['T1']}}),
        ]

        a = student('A',
                    avail=[('月', '16')],
                    backup_avail=[('火', '17')],
                    ng_students=['B'],
                    needs={'数': 2})
        b = student('B',
                    avail=[('月', '16'), ('火', '17')],
                    wish_teachers=['T1'],
                    needs={'数': 2})

        schedule, unplaced, _ = build_schedule([a, b], wt, SKILLS, {}, {})

        assert unplaced == [], f"Unexpected unplaced: {unplaced}"

        a_places = find_placements(schedule, 'A')
        assert len(a_places) == 2, f"A should have 2 placements, got {a_places}"

        for wi, day, ts in a_places:
            assert is_primary(a, day, ts), \
                f"Week {wi}: A at ({day},{ts}) should be primary"


# ---------------------------------------------------------------------------
# Test 5: backup_time 重み変更 → backup_slot 使用減少
# ---------------------------------------------------------------------------

class TestBackupTimeWeight:
    """Test 5: backup_time 重み変更 — primary vs backup+continuous の選択を検証"""

    def _make_inputs(self):
        """
        T1 at Mon 16, Mon 17, Tue 17.
        Student A has a FIXED session at Mon 16 for 数 (→ existing on Mon).
        A's 2nd subject (英) must choose between:
          - Mon 17 (backup, but continuous with Mon 16 → +2000 bonus)
          - Tue 17 (primary, no extra bonus)
        """
        wt = [week({'月': {'16': ['T1'], '17': ['T1']}, '火': {'17': ['T1']}})]
        a = student('A',
                    avail=[('月', '16'), ('火', '17')],
                    backup_avail=[('月', '17')],
                    needs={'数': 1, '英': 1},
                    fixed=[('月', '16', '数')])
        return a, wt

    def test_old_weight_prefers_backup_continuous(self):
        """
        With backup_time=-150 (old):
        Mon 17 score = -150 + 2000 = 1850 > Tue 17 score = 0.
        → 英 is placed at BACKUP Mon 17.
        """
        a, wt = self._make_inputs()
        old_weights = {**DEFAULT_WEIGHTS, 'backup_time': -150}
        schedule, unplaced, _ = build_schedule([a], wt, SKILLS, {}, {},
                                               weights=old_weights)

        assert unplaced == []
        places = find_placements(schedule, 'A')
        assert len(places) == 2

        # The 英 placement is the one that is NOT the fixed Mon 16 slot
        english = [(d, ts) for _, d, ts in places if (d, ts) != ('月', '16')]
        assert len(english) == 1
        day_e, ts_e = english[0]

        assert (day_e, ts_e) == ('月', '17'), \
            f"Old weights: expected backup Mon 17, got ({day_e}, {ts_e})"

    def test_new_weight_prefers_primary_over_backup_continuous(self):
        """
        With backup_time=-2100 (new / DEFAULT_WEIGHTS):
        Mon 17 score = -2100 + 2000 = -100 < Tue 17 score = 0.
        → 英 is placed at PRIMARY Tue 17.
        """
        a, wt = self._make_inputs()
        schedule, unplaced, _ = build_schedule([a], wt, SKILLS, {}, {})

        assert unplaced == []
        places = find_placements(schedule, 'A')
        assert len(places) == 2

        english = [(d, ts) for _, d, ts in places if (d, ts) != ('月', '16')]
        assert len(english) == 1
        day_e, ts_e = english[0]

        assert (day_e, ts_e) == ('火', '17'), \
            f"New weights: expected primary Tue 17, got ({day_e}, {ts_e})"
