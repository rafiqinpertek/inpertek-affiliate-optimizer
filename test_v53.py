"""Uji algoritma V5.3 "Margin Berjarak Aman". Jalankan: python test_v53.py"""
import math, random
from engine import analyze, MODES

def rec(price, old_pct, margin_before):
    """margin_before = margin bersih master + komisi awal."""
    fee = price * old_pct / 100
    return {'validation': 'OK', 'price': price, 'old_pct': old_pct, 'old_fee': fee, 'old_margin': margin_before - fee}

def check(r, kw):
    out = analyze(r, **kw)
    k = MODES[kw.get('mode', 'AGRESIF')]
    assert out['status'] in ('NAIK', 'TURUNKAN', 'TETAP', 'REVIEW NONAKTIF'), out
    before = r['old_margin'] + r['old_fee']; old = r['old_pct']; price = r['price']
    if out['status'] == 'REVIEW NONAKTIF':
        assert before <= 0 or before < kw.get('min_profit', 0) or math.floor(out['hard_cap'] + 1e-9) < 1, (r, out)
        return out
    p = out['proposed_pct']; cost = price * p / 100; remain = before - cost
    assert cost > 0, ('komisi usulan nol', r, out)
    assert remain > cost, ('komisi tidak lebih kecil dari sisa margin', r, out)
    assert remain + 0.01 >= k * cost, ('jarak aman dilanggar', r, out)
    assert remain + 0.01 >= kw.get('min_profit', 0), ('di bawah margin minimum', r, out)
    cap_i = math.floor(out['hard_cap'] + 1e-9)
    if out['status'] == 'TETAP':
        assert p == old and cap_i <= old <= out['hard_cap'] + 1e-9, (r, out)
    else:
        assert p == int(p) and p <= min(kw.get('max_pct', 30), kw.get('target_pct', 30)) + 1e-9, (r, out)
    if out['status'] == 'TURUNKAN':
        assert old > out['hard_cap'] and p == cap_i, ('turun tidak pas ke batas', r, out)
    if out['status'] == 'NAIK':
        assert p > old and p <= cap_i, (r, out)
        if kw.get('max_increase') is None:
            assert p == cap_i, ('naik tidak pas ke batas', r, out)
        else:
            assert p <= old + kw['max_increase'] + 1e-9, (r, out)
    return out

def test_random(n=60000):
    random.seed(53)
    for _ in range(n):
        price = random.choice([5000, 15000, 25000, 49900, 99000, 250000, 1_200_000])
        old = random.choice([0, 1, 2, 2.5, 3, 4.7, 5, 8, 10, 12, 15, 20, 25, 30])
        kw = dict(mode=random.choice(list(MODES)), min_profit=random.choice([0, 0, 1000, 5000]),
                  max_pct=random.choice([10, 20, 30]), target_pct=30, max_increase=random.choice([None, None, 1, 3, 5]))
        check(rec(price, old, price * random.uniform(-0.1, 0.9)), kw)

def test_examples():
    # Harga Rp100.000, margin bersih Rp20.000 + komisi awal 5% = Rp25.000 sebelum affiliate.
    # Agresif (1,25x): batas = 25.000/2,25 = Rp11.111 -> 11%. Sisa Rp14.000 = 1,27x komisi.
    out = analyze(rec(100000, 5, 25000))
    assert out['status'] == 'NAIK' and out['proposed_pct'] == 11
    assert analyze(rec(100000, 5, 25000), mode='SEIMBANG')['proposed_pct'] == 10      # 25.000/2,5
    assert analyze(rec(100000, 5, 25000), mode='KONSERVATIF')['proposed_pct'] == 8    # 25.000/3
    # Komisi awal 12% dengan margin sebelum affiliate Rp25.000 -> sisa Rp13.000 (1,08x) terlalu mepet -> turun ke 11%
    out = analyze(rec(100000, 12, 25000))
    assert out['status'] == 'TURUNKAN' and out['proposed_pct'] == 11
    # Sudah pas
    assert analyze(rec(100000, 11, 25000))['status'] == 'TETAP'
    # Margin Rp2.000 untuk harga Rp100.000: komisi 1% (Rp1.000) menyisakan Rp1.000 = 1x < 1,25x -> nonaktif
    assert analyze(rec(100000, 3, 2000))['status'] == 'REVIEW NONAKTIF'

if __name__ == '__main__':
    test_examples(); test_random()
    print('Semua tes lulus.')
