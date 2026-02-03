import math
import datetime
import statistics
import re
from datetime import timedelta

# Excel Error Constants
ERR_NULL = "#NULL!"
ERR_DIV0 = "#DIV/0!"
ERR_VALUE = "#VALUE!"
ERR_REF = "#REF!"
ERR_NAME = "#NAME?"
ERR_NUM = "#NUM!"
ERR_NA = "#N/A"
ERROR_VALUES = {ERR_NULL, ERR_DIV0, ERR_VALUE, ERR_REF, ERR_NAME, ERR_NUM, ERR_NA}

# --- Helpers ---
def is_error(val):
    return isinstance(val, str) and val in ERROR_VALUES

def flat_list(args):
    """Recursively flattens lists/grids into a single flat list of values."""
    res = []
    for arg in args:
        if isinstance(arg, list):
            res.extend(flat_list(arg))
        else:
            res.append(arg)
    return res

def to_num(val):
    """Coerces value to float, propagating Excel errors."""
    if is_error(val): return val
    if val is None or val == "": return 0.0
    if isinstance(val, bool): return 1.0 if val else 0.0
    if isinstance(val, (int, float)): return float(val)
    if isinstance(val, (datetime.date, datetime.datetime)):
        return to_excel_date(val)
    try:
        s = str(val).strip()
        if not s: return 0.0
        if s.upper() == "TRUE": return 1.0
        if s.upper() == "FALSE": return 0.0
        return float(s.replace(",", ".")) if "." not in s or "," not in s else float(s)
    except (ValueError, TypeError):
        return ERR_VALUE

def _match_criteria(val, criterion):
    if is_error(val): return False
    if not isinstance(criterion, str):
        return val == criterion
    match = re.match(r'^(>=|<=|<>|>|<|=)?(.*)$', criterion)
    op, target = match.groups()
    if op is None: op = "="
    try:
        t_num = float(target)
        v_num = to_num(val)
        if is_error(v_num): return False
        if op == "=": return v_num == t_num
        if op == "<>": return v_num != t_num
        if op == ">": return v_num > t_num
        if op == "<": return v_num < t_num
        if op == ">=": return v_num >= t_num
        if op == "<=": return v_num <= t_num
    except:
        v_str, t_str = str(val).lower(), target.lower()
        if op == "=":
            pattern = "^" + re.escape(t_str).replace(r'\*', '.*').replace(r'\?', '.') + "$"
            return bool(re.match(pattern, v_str))
        if op == "<>": return v_str != t_str
    return False

def _parse_addr(addr):
    if not addr: return "A", 1
    m = re.match(r"([A-Z]+)([0-9]+)", str(addr).upper().replace("$", ""))
    if not m: return "A", 1
    return m.group(1), int(m.group(2))

def _col_to_int(col):
    num = 0
    for char in col.upper(): num = num * 26 + (ord(char) - ord('A') + 1)
    return num

def _int_to_col(n):
    name = ""
    while n > 0: n, rem = divmod(n - 1, 26); name = chr(65 + rem) + name
    return name

# Math Functions
def excel_ABS(v): n = to_num(v); return n if is_error(n) else abs(n)
def excel_SUM(*args):
    total = 0.0; flat = flat_list(args)
    for x in flat:
        if is_error(x): return x
        v = to_num(x); total += v if isinstance(v, float) else 0.0
    return total
def excel_PRODUCT(*args):
    prod = 1.0; flat = flat_list(args)
    if not flat: return 0.0
    for x in flat:
        if is_error(x): return x
        v = to_num(x); prod *= v if isinstance(v, float) else 1.0
    return prod
def excel_QUOTIENT(n, d): vn, vd = to_num(n), to_num(d); return vn if is_error(vn) else (vd if is_error(vd) else (ERR_DIV0 if vd == 0 else float(int(vn / vd))))
def excel_MOD(n, d): vn, vd = to_num(n), to_num(d); return vn if is_error(vn) else (vd if is_error(vd) else (ERR_DIV0 if vd == 0 else vn % vd))
def excel_POWER(b, e): vb, ve = to_num(b), to_num(e); return vb if is_error(vb) else (ve if is_error(ve) else (vb ** ve if vb != 0 or ve >= 0 else ERR_NUM))
def excel_SQRT(n): vn = to_num(n); return vn if is_error(vn) else (ERR_NUM if vn < 0 else math.sqrt(vn))
def excel_EXP(n): vn = to_num(n); return vn if is_error(vn) else math.exp(vn)
def excel_LN(n): vn = to_num(n); return vn if is_error(vn) else (math.log(vn) if vn > 0 else ERR_NUM)
def excel_LOG10(n): vn = to_num(n); return vn if is_error(vn) else (math.log10(vn) if vn > 0 else ERR_NUM)
def excel_LOG(n, b=10): vn, vb = to_num(n), to_num(b); return vn if is_error(vn) else (vb if is_error(vb) else (math.log(vn, vb) if vn > 0 and vb > 0 and vb != 1 else ERR_NUM))
def excel_INT(n): v = to_num(n); return v if is_error(v) else float(math.floor(v))
def excel_TRUNC(n, d=0): vn, vd = to_num(n), to_num(d); f = 10**int(vd); return float(int(vn * f))/f if not is_error(vn) else vn
def excel_ROUND(n, d=0): vn, vd = to_num(n), to_num(d); return round(vn, int(vd)) if not is_error(vn) else vn
def excel_ROUNDUP(n, d=0): vn, vd = to_num(n), to_num(d); f = 10**int(vd); return (math.ceil(vn * f) / f if vn > 0 else math.floor(vn * f) / f) if not is_error(vn) else vn
def excel_ROUNDDOWN(n, d=0): vn, vd = to_num(n), to_num(d); f = 10**int(vd); return (math.floor(vn * f) / f if vn > 0 else math.ceil(vn * f) / f) if not is_error(vn) else vn
def excel_SIGN(n): v = to_num(n); return v if is_error(v) else (1.0 if v > 0 else (-1.0 if v < 0 else 0.0))
def excel_PI(): return math.pi
def excel_DEGREES(v): n = to_num(v); return n if is_error(n) else n * 180.0 / math.pi
def excel_RADIANS(v): n = to_num(v); return n if is_error(n) else n * math.pi / 180.0

# Trig
def excel_SIN(n): v = to_num(n); return v if is_error(v) else math.sin(v)
def excel_COS(n): v = to_num(n); return v if is_error(v) else math.cos(v)
def excel_TAN(n): v = to_num(n); return v if is_error(v) else math.tan(v)
def excel_ASIN(n): v = to_num(n); return v if is_error(v) else (math.asin(v) if -1 <= v <= 1 else ERR_NUM)
def excel_ACOS(n): v = to_num(n); return v if is_error(v) else (math.acos(v) if -1 <= v <= 1 else ERR_NUM)
def excel_ATAN(n): v = to_num(n); return v if is_error(v) else math.atan(v)
def excel_ATAN2(x, y): vx, vy = to_num(x), to_num(y); return vx if is_error(vx) else (vy if is_error(vy) else (math.atan2(vy, vx) if vx != 0 or vy != 0 else ERR_DIV0))
def excel_SINH(n): v = to_num(n); return v if is_error(v) else math.sinh(v)
def excel_COSH(n): v = to_num(n); return v if is_error(v) else math.cosh(v)
def excel_TANH(n): v = to_num(n); return v if is_error(v) else math.tanh(v)
def excel_ASINH(n): v = to_num(n); return v if is_error(v) else math.asinh(v)
def excel_ACOSH(n): v = to_num(n); return v if is_error(v) else (math.acosh(v) if v >= 1 else ERR_NUM)
def excel_ATANH(n): v = to_num(n); return v if is_error(v) else (math.atanh(v) if -1 < v < 1 else ERR_NUM)
def excel_CSC(n): v = to_num(n); return v if is_error(v) else (1.0/math.sin(v) if math.sin(v) != 0 else ERR_DIV0)
def excel_SEC(n): v = to_num(n); return v if is_error(v) else (1.0/math.cos(v) if math.cos(v) != 0 else ERR_DIV0)
def excel_COT(n): v = to_num(n); return v if is_error(v) else (1.0/math.tan(v) if math.tan(v) != 0 else ERR_DIV0)
def excel_CSCH(n): v = to_num(n); return v if is_error(v) else (1.0/math.sinh(v) if math.sinh(v) != 0 else ERR_DIV0)
def excel_SECH(n): v = to_num(n); return v if is_error(v) else (1.0/math.cosh(v))
def excel_COTH(n): v = to_num(n); return v if is_error(v) else (1.0/math.tanh(v) if math.tanh(v) != 0 else ERR_DIV0)
def excel_ACOT(n): v = to_num(n); return v if is_error(v) else (math.atan(1/v) if v != 0 else math.pi/2)
def excel_ACOTH(n): v = to_num(n); return v if is_error(v) else (0.5 * math.log((v + 1) / (v - 1)) if abs(v) > 1 else ERR_NUM)

# GCD/LCM
def excel_GCD(*args):
    nums = [int(to_num(x)) for x in flat_list(args) if not is_error(to_num(x))]
    if any(n < 0 for n in nums): return ERR_NUM
    if not nums: return 0.0
    res = nums[0]
    for n in nums[1:]: res = math.gcd(res, n)
    return float(res)
def excel_LCM(*args):
    nums = [int(to_num(x)) for x in flat_list(args) if not is_error(to_num(x))]
    if any(n < 0 for n in nums): return ERR_NUM
    if not nums: return 1.0
    res = nums[0]
    for n in nums[1:]: res = abs(res * n) // math.gcd(res, n) if res and n else 0
    return float(res)

# Rounding
def excel_CEILING(v, s=1): 
    nv, ns = to_num(v), to_num(s)
    if is_error(nv): return nv
    if is_error(ns): return ns
    if ns == 0: return 0.0
    if nv > 0 and ns < 0: return ERR_NUM
    return math.ceil(nv/ns)*ns
def excel_FLOOR(v, s=1):
    nv, ns = to_num(v), to_num(s)
    if is_error(nv): return nv
    if is_error(ns): return ns
    if ns == 0: return ERR_DIV0
    if (nv >= 0 and ns < 0) or (nv < 0 and ns > 0): return ERR_NUM
    return math.floor(nv/ns)*ns
def excel_MROUND(v, s):
    nv, ns = to_num(v), to_num(s)
    if is_error(nv) or is_error(ns): return nv if is_error(nv) else ns
    if ns == 0: return 0.0
    if (nv > 0 and ns < 0) or (nv < 0 and ns > 0): return ERR_NUM
    return round(nv / ns) * ns
def excel_EVEN(v):
    nv = to_num(v)
    if is_error(nv): return nv
    if nv == 0: return 0.0
    a = math.ceil(abs(nv)); res = a if a % 2 == 0 else a + 1
    return float(res if nv > 0 else -res)
def excel_ODD(v):
    nv = to_num(v)
    if is_error(nv): return nv
    if nv == 0: return 1.0
    a = math.ceil(abs(nv)); res = a if a % 2 != 0 else a + 1
    return float(res if nv > 0 else -res)
def excel_CEILING_MATH(v, s=1, m=0): return excel_CEILING(v, s)
def excel_FLOOR_MATH(v, s=1, m=0): return excel_FLOOR(v, s)
def excel_CEILING_PRECISE(v, s=1): return excel_CEILING(v, s)
def excel_FLOOR_PRECISE(v, s=1): return excel_FLOOR(v, s)

# Base Conversion
def excel_BASE(n, r, l=0):
    n, r, l = int(to_num(n)), int(to_num(r)), int(to_num(l))
    if not (2 <= r <= 36): return ERR_NUM
    res = ""; digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    while n > 0: n, rem = divmod(n, r); res = digits[rem] + res
    return res.zfill(l) if res else "0".zfill(l)
def excel_DECIMAL(t, r):
    try: return float(int(str(t), int(to_num(r))))
    except: return ERR_NUM
def excel_BIN2DEC(s): s = str(s); return float(int(s, 2)) if len(s) < 10 else float(int(s, 2) - 1024)
def excel_DEC2BIN(n, p=None):
    n = int(to_num(n)); b = bin(n & 0x3FF)[2:]
    return b.zfill(int(p)) if p else b
def excel_DEC2HEX(n, p=None):
    n = int(to_num(n)); h = hex(n & 0xFFFFFFFFFF)[2:].upper()
    return h.zfill(int(p)) if p else h
def excel_HEX2DEC(s):
    try: s = str(s); return float(int(s, 16)) if len(s) < 10 else float(int(s, 16) - 0x10000000000)
    except: return ERR_NUM
def excel_BIN2HEX(s, p=None): return excel_DEC2HEX(excel_BIN2DEC(s), p)
def excel_BIN2OCT(s, p=None): return excel_DEC2OCT(excel_BIN2DEC(s), p)
def excel_DEC2OCT(n, p=None): 
    n = int(to_num(n)); o = oct(n & 0x3FFFFFFF)[2:]
    return o.zfill(int(p)) if p else o
def excel_HEX2BIN(s, p=None): return excel_DEC2BIN(excel_HEX2DEC(s), p)
def excel_HEX2OCT(s, p=None): return excel_OCT2OCT(excel_HEX2DEC(s), p)
def excel_OCT2BIN(s, p=None): return excel_DEC2BIN(excel_OCT2DEC(s), p)
def excel_OCT2DEC(s):
    try: s = str(s); return float(int(s, 8)) if len(s) < 10 else float(int(s, 8) - 0x40000000)
    except: return ERR_NUM
def excel_OCT2HEX(s, p=None): return excel_DEC2HEX(excel_OCT2DEC(s), p)

# Logical & Info
def excel_AND(*args): flat = flat_list(args); return ERR_VALUE if not flat else (next((a for a in flat if is_error(a)), None) or all(to_num(x) for x in flat))
def excel_OR(*args): flat = flat_list(args); return ERR_VALUE if not flat else (next((a for a in flat if is_error(a)), None) or any(to_num(x) for x in flat))
def excel_NOT(x): v = to_num(x); return v if is_error(v) else not bool(v)
def excel_XOR(*args): flat = [bool(to_num(a)) for a in flat_list(args) if not is_error(to_num(a))]; return sum(flat) % 2 != 0
def excel_IF(c, t, f=False): return t if c else f
def excel_IFERROR(v, e): return e if is_error(v) else v
def excel_IFNA(v, e): return e if v == ERR_NA else v
def excel_IFS(*args):
    for i in range(0, len(args), 2):
        if to_num(args[i]): return args[i+1]
    return ERR_NA
def excel_SWITCH(v, *args):
    for i in range(0, len(args)-1, 2):
        if v == args[i]: return args[i+1]
    return args[-1] if len(args) % 2 != 0 else ERR_NA
def excel_N(v):
    if is_error(v): return v
    if isinstance(v, bool): return 1.0 if v else 0.0
    if isinstance(v, str): return 0.0
    return to_num(v)
def excel_TYPE(v):
    if is_error(v): return 16.0
    if isinstance(v, bool): return 4.0
    if isinstance(v, (int, float)): return 1.0
    if isinstance(v, str): return 2.0
    if isinstance(v, list): return 64.0
    return 1.0
def excel_ERROR_TYPE(v):
    m = {ERR_NULL: 1, ERR_DIV0: 2, ERR_VALUE: 3, ERR_REF: 4, ERR_NAME: 5, ERR_NUM: 6, ERR_NA: 7}
    return float(m.get(v, ERR_NA))
def excel_ISERR(v): return is_error(v) and v != ERR_NA
def excel_ISERROR(v): return is_error(v)
def excel_ISNA(v): return v == ERR_NA
def excel_ISLOGICAL(v):
    if is_error(v): return v
    return isinstance(v, bool)
def excel_ISNUMBER(v):
    if is_error(v): return v
    return isinstance(v, (int, float))
def excel_ISTEXT(v):
    if is_error(v): return v
    return isinstance(v, str)
def excel_ISNONTEXT(v):
    if is_error(v): return v
    return not isinstance(v, str)
def excel_ISBLANK(v): return v is None or v == ""
def excel_ISEVEN(v):
    n = to_num(v)
    if is_error(n): return n
    return float(int(n) % 2 == 0)
def excel_ISODD(v):
    n = to_num(v)
    if is_error(n): return n
    return float(int(n) % 2 != 0)
def excel_ISREF(v): return isinstance(v, dict) and v.get("type") in ("cell_reference", "range_reference")

# Statistical
def excel_AVERAGE(*args):
    flat = flat_list(args); nums = [to_num(x) for x in flat if isinstance(to_num(x), float) and not is_error(to_num(x))]
    return sum(nums)/len(nums) if nums else ERR_DIV0
def excel_AVERAGEA(*args):
    flat = flat_list(args); res = [to_num(x) for x in flat if not is_error(to_num(x))]
    return sum(res)/len(res) if res else ERR_DIV0
def excel_COUNT(*args):
    return float(sum(1 for x in flat_list(args) if isinstance(to_num(x), float) and not is_error(to_num(x))))
def excel_COUNTA(*args):
    return float(len([x for x in flat_list(args) if x is not None and x != ""]))
def excel_MAX(*args):
    flat = flat_list(args); res = [to_num(x) for x in flat if not is_error(to_num(x))]
    return max(res) if res else 0.0
def excel_MAXA(*args):
    flat = flat_list(args); res = [to_num(x) for x in flat if not is_error(to_num(x))]
    return max(res) if res else 0.0
def excel_MIN(*args):
    flat = flat_list(args); res = [to_num(x) for x in flat if not is_error(to_num(x))]
    return min(res) if res else 0.0
def excel_MINA(*args):
    flat = flat_list(args); res = [to_num(x) for x in flat if not is_error(to_num(x))]
    return min(res) if res else 0.0
def excel_MEDIAN(*args):
    flat = flat_list(args); nums = sorted([to_num(x) for x in flat if not is_error(to_num(x))])
    return statistics.median(nums) if nums else ERR_NUM
def excel_STDEV(*args):
    flat = flat_list(args); nums = [to_num(x) for x in flat if isinstance(to_num(x), float) and not is_error(to_num(x))]
    return statistics.stdev(nums) if len(nums) > 1 else ERR_DIV0
def excel_STDEVP(*args):
    flat = flat_list(args); nums = [to_num(x) for x in flat if isinstance(to_num(x), float) and not is_error(to_num(x))]
    return statistics.pstdev(nums) if nums else ERR_DIV0
def excel_VAR(*args):
    flat = flat_list(args); nums = [to_num(x) for x in flat if isinstance(to_num(x), float) and not is_error(to_num(x))]
    return statistics.variance(nums) if len(nums) > 1 else ERR_DIV0
def excel_VARP(*args):
    flat = flat_list(args); nums = [to_num(x) for x in flat if isinstance(to_num(x), float) and not is_error(to_num(x))]
    return statistics.pvariance(nums) if nums else ERR_DIV0
def excel_RANK(n, r, o=0):
    nv = to_num(n)
    if is_error(nv): return nv
    flat = flat_list(r); nums = sorted([to_num(x) for x in flat if not is_error(to_num(x))], reverse=(to_num(o) == 0))
    return float(nums.index(nv)+1) if nv in nums else ERR_NA
def excel_LARGE(a, k):
    kv = to_num(k)
    if is_error(kv): return kv
    flat = flat_list(a); nums = sorted([to_num(x) for x in flat if not is_error(to_num(x))], reverse=True)
    return nums[int(kv)-1] if 1 <= int(kv) <= len(nums) else ERR_NUM
def excel_SMALL(a, k):
    kv = to_num(k)
    if is_error(kv): return kv
    flat = flat_list(a); nums = sorted([to_num(x) for x in flat if not is_error(to_num(x))])
    return nums[int(kv)-1] if 1 <= int(kv) <= len(nums) else ERR_NUM
def excel_PERCENTILE(a, k):
    kv = to_num(k)
    if is_error(kv): return kv
    flat = flat_list(a); nums = sorted([to_num(x) for x in flat if not is_error(to_num(x))])
    if not (0 <= kv <= 1) or not nums: return ERR_NUM
    idx = kv * (len(nums)-1); i, f = int(idx), idx-int(idx)
    return nums[i] + (nums[i+1]-nums[i])*f if i < len(nums)-1 else nums[-1]
def excel_COUNTIF(rng, criteria): return float(sum(1 for val in flat_list([rng]) if _match_criteria(val, criteria)))
def excel_SUMIF(rng, criteria, sum_rng=None):
    if sum_rng is None: sum_rng = rng
    s_vals, c_vals = flat_list([sum_rng]), flat_list([rng]); total = 0.0
    for i in range(min(len(c_vals), len(s_vals))):
        if _match_criteria(c_vals[i], criteria): total += to_num(s_vals[i]) if not is_error(to_num(s_vals[i])) else 0.0
    return total
def excel_AVERAGEIF(rng, criteria, avg_rng=None):
    if avg_rng is None: avg_rng = rng
    a_vals, c_vals = flat_list([avg_rng]), flat_list([rng]); total, count = 0.0, 0
    for i in range(min(len(c_vals), len(a_vals))):
        if _match_criteria(c_vals[i], criteria): total += to_num(a_vals[i]); count += 1
    return total / count if count > 0 else ERR_DIV0
def excel_COUNTIFS(*args):
    ranges = [flat_list([args[i]]) for i in range(0, len(args), 2)]; criteria = [args[i+1] for i in range(0, len(args), 2)]
    return float(sum(1 for i in range(len(ranges[0])) if all(_match_criteria(ranges[j][i], criteria[j]) for j in range(len(ranges)))))
def excel_SUMIFS(sum_rng, *args):
    s_vals = flat_list([sum_rng]); ranges = [flat_list([args[i]]) for i in range(0, len(args), 2)]; crit = [args[i+1] for i in range(0, len(args), 2)]
    return sum(to_num(s_vals[i]) for i in range(len(s_vals)) if all(_match_criteria(ranges[j][i], crit[j]) for j in range(len(ranges))))
def excel_MAXIFS(max_rng, *args):
    m_vals = flat_list([max_rng]); ranges = [flat_list([args[i]]) for i in range(0, len(args), 2)]; crit = [args[i+1] for i in range(0, len(args), 2)]
    res = [to_num(m_vals[i]) for i in range(len(m_vals)) if all(_match_criteria(ranges[j][i], crit[j]) for j in range(len(ranges)))]
    return max(res) if res else 0.0
def excel_MINIFS(min_rng, *args):
    m_vals = flat_list([min_rng]); ranges = [flat_list([args[i]]) for i in range(0, len(args), 2)]; crit = [args[i+1] for i in range(0, len(args), 2)]
    res = [to_num(m_vals[i]) for i in range(len(m_vals)) if all(_match_criteria(ranges[j][i], crit[j]) for j in range(len(ranges)))]
    return min(res) if res else 0.0
def excel_COUNTBLANK(rng): return float(sum(1 for x in flat_list([rng]) if x is None or x == ""))

# Text
def excel_CHAR(n): v = to_num(n); return v if is_error(v) else (chr(int(v)) if 0 <= v <= 1114111 else ERR_VALUE)
def excel_CODE(s): return s if is_error(s) else (ord(str(s)[0]) if s else ERR_VALUE)
def excel_CLEAN(s): return s if is_error(s) else "".join(c for c in str(s) if ord(c) >= 32)
def excel_REPT(s, n): v = to_num(n); return str(s) * int(v) if not is_error(v) else v
def excel_EXACT(s1, s2): return str(s1) == str(s2)
def excel_CONCAT(*args):
    for a in flat_list(args):
        if is_error(a): return a
    return "".join(str(x) for x in flat_list(args))
def excel_CONCATENATE(*args): return excel_CONCAT(*args)
def excel_LEFT(s, n=1): return str(s)[:max(0, int(to_num(n)))]
def excel_RIGHT(s, n=1): v = int(to_num(n)); return str(s)[-v:] if v > 0 else ""
def excel_MID(s, start, n): start, n = int(to_num(start)), int(to_num(n)); return str(s)[max(0, start-1):max(0, start-1)+n] if start > 0 else ERR_VALUE
def excel_LEN(s): return float(len(str(s)))
def excel_UPPER(s): return str(s).upper()
def excel_LOWER(s): return str(s).lower()
def excel_PROPER(s): return str(s).title()
def excel_TRIM(s): return " ".join(str(s).split())
def excel_VALUE(s): return to_num(s)
def excel_T(s): return s if isinstance(s, str) else ""
def excel_FIND(f, w, n=1): n = int(to_num(n)); idx = str(w).find(str(f), n-1); return float(idx+1) if idx >= 0 else ERR_VALUE
def excel_SEARCH(f, w, n=1): n = int(to_num(n)); f_pat = re.escape(str(f).lower()).replace(r'\*', '.*').replace(r'\?', '.'); m = re.search(f_pat, str(w).lower()[n-1:]); return float(m.start() + n) if m else ERR_VALUE
def excel_SUBSTITUTE(s, o, n, i=None):
    s, o, n = str(s), str(o), str(n)
    if i is None: return s.replace(o, n)
    i = int(to_num(i)); parts = s.split(o); return o.join(parts[:i]) + n + o.join(parts[i:]) if 1 <= i < len(parts) else s
def excel_REPLACE(o, s, n, w): s, n = int(to_num(s)), int(to_num(n)); return str(o)[:s-1] + str(w) + str(o)[s-1+n:]
def excel_TEXT(v, f): return str(v)

# Matrix
def excel_MMULT(a1, a2):
    r1, c1, r2, c2 = len(a1), len(a1[0]), len(a2), len(a2[0])
    if c1 != r2: return ERR_VALUE
    res = [[0.0]*c2 for _ in range(r1)]
    for i in range(r1):
        for j in range(c2):
            for k in range(c1):
                v1, v2 = to_num(a1[i][k]), to_num(a2[k][j])
                if is_error(v1) or is_error(v2): return v1 if is_error(v1) else v2
                res[i][j] += v1 * v2
    return res
def excel_TRANSPOSE(a): return [[a[j][i] for j in range(len(a))] for i in range(len(a[0]))] if isinstance(a, list) and isinstance(a[0], list) else a
def excel_MDETERM(m):
    if not isinstance(m, list) or not isinstance(m[0], list) or len(m) != len(m[0]): return ERR_VALUE
    n = len(m); v = [[to_num(x) for x in row] for row in m]; det = 1
    for i in range(n):
        pivot = i
        for k in range(i+1, n):
            if abs(v[k][i]) > abs(v[pivot][i]): pivot = k
        v[i], v[pivot] = v[pivot], v[i]
        if i != pivot: det *= -1
        if v[i][i] == 0: return 0.0
        det *= v[i][i]
        for k in range(i+1, n):
            factor = v[k][i] / v[i][i]
            for j in range(i+1, n): v[k][j] -= factor * v[i][j]
    return det
def excel_MINVERSE(m):
    if not isinstance(m, list) or not isinstance(m[0], list) or len(m) != len(m[0]): return ERR_VALUE
    n = len(m); aug = [[to_num(m[i][j]) for j in range(n)] + [1 if i == k else 0 for k in range(n)] for i in range(n)]
    for i in range(n):
        pivot = i
        for k in range(i+1, n):
            if abs(aug[k][i]) > abs(aug[pivot][i]): pivot = k
        aug[i], aug[pivot] = aug[pivot], aug[i]
        if aug[i][i] == 0: return ERR_NUM
        f = aug[i][i]
        for j in range(i, 2*n): aug[i][j] /= f
        for k in range(n):
            if k != i:
                f = aug[k][i]
                for j in range(i, 2*n): aug[k][j] -= f * aug[i][j]
    return [row[n:] for row in aug]

# Financial
def excel_PV(r, n, p, f=0, t=0): r, n, p, f, t = to_num(r), to_num(n), to_num(p), to_num(f), to_num(t); return (-(f + p * n) if r == 0 else -(f + p * (1 + r * t) * ((pow(1 + r, n) - 1) / r)) / pow(1 + r, n))
def excel_FV(r, n, p, v=0, t=0): r, n, p, v, t = to_num(r), to_num(n), to_num(p), to_num(v), to_num(t); return (-(v + p * n) if r == 0 else -v * pow(1 + r, n) - p * (1 + r * t) * ((pow(1 + r, n) - 1) / r))
def excel_PMT(r, n, v, f=0, t=0): r, n, v, f, t = to_num(r), to_num(n), to_num(v), to_num(f), to_num(t); return (-(v + f) / n if r == 0 else -(v * pow(1 + r, n) + f) / ((1 + r * t) * (pow(1 + r, n) - 1) / r))
def excel_NPER(r, p, v, f=0, t=0):
    r, p, v, f, t = to_num(r), to_num(p), to_num(v), to_num(f), to_num(t)
    if r == 0: return -(v + f) / p
    try: return math.log((p*(1+r*t)/r - f)/(v + p*(1+r*t)/r)) / math.log(1+r)
    except: return ERR_NUM
def excel_NPV(rate, *values): r = to_num(rate); return sum(to_num(v) / pow(1 + r, i + 1) for i, v in enumerate(flat_list(values)))
def excel_IRR(values, guess=0.1):
    vals = [to_num(v) for v in flat_list(values)]; rate = guess
    for _ in range(100):
        npv = sum(v / pow(1 + rate, i) for i, v in enumerate(vals))
        if abs(npv) < 1e-7: return rate
        d_npv = sum(-i * v / pow(1 + rate, i + 1) for i, v in enumerate(vals))
        if d_npv == 0: break
        rate -= npv / d_npv
    return ERR_NUM
def excel_RATE(nper, pmt, pv, fv=0, type=0, guess=0.1):
    n, p, v, f, t = to_num(nper), to_num(pmt), to_num(pv), to_num(fv), to_num(type); rate = guess
    for _ in range(100):
        val = v*(1+rate)**n + p*(1+rate*t)*((pow(1+rate,n)-1)/rate) + f if rate != 0 else v+p*n+f
        if abs(val) < 1e-7: return rate
        d_val = (v*n*(1+rate)**(n-1) + p*(1+rate*t)*(n*(1+rate)**(n-1)/rate - (pow(1+rate,n)-1)/rate**2)) if rate != 0 else n*v + p*n*(n-1)/2
        if d_val == 0: break
        rate -= val / d_val
    return rate
def excel_LOG(n, b=10):
    vn, vb = to_num(n), to_num(b)
    if is_error(vn): return vn
    if is_error(vb): return vb
    if vb == 1: return ERR_DIV0
    if vn <= 0 or vb <= 0: return ERR_NUM
    return math.log(vn, vb)

# Date/Time
BASE_DATE = datetime.datetime(1899, 12, 30)
def to_excel_date(dt): val = (dt - BASE_DATE if isinstance(dt, datetime.datetime) else datetime.datetime.combine(dt, datetime.time()) - BASE_DATE); return float(val.days) + (val.seconds / 86400.0)
def _get_dt(s): v = to_num(s); return v if is_error(v) else (datetime.datetime(1899,12,31)+timedelta(days=int(v)) if v < 60 else BASE_DATE+timedelta(days=int(v)))
def excel_DATE(y, m, d):
    y, m, d = int(to_num(y)), int(to_num(m)), int(to_num(d))
    if y < 1900: y += 1900
    try:
        dt = datetime.datetime(y, 1, 1) + timedelta(days=d-1); tm = m - 1; dt = dt.replace(year=dt.year + tm // 12, month=(tm % 12) + 1)
        return to_excel_date(dt) + (1 if (dt - BASE_DATE).days >= 60 else 0)
    except: return ERR_VALUE
def excel_YEAR(s): dt = _get_dt(s); return float(dt.year) if not is_error(dt) else dt
def excel_MONTH(s): dt = _get_dt(s); return float(dt.month) if not is_error(dt) else dt
def excel_DAY(s): dt = _get_dt(s); return float(dt.day) if not is_error(dt) else dt
def excel_HOUR(s): dt = _get_dt(s); return float(dt.hour) if not is_error(dt) else dt
def excel_MINUTE(s): dt = _get_dt(s); return float(dt.minute) if not is_error(dt) else dt
def excel_SECOND(s): dt = _get_dt(s); return float(dt.second) if not is_error(dt) else dt
def excel_DATEVALUE(s):
    s = str(s).strip()
    for f in ("%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y"):
        try: return to_excel_date(datetime.datetime.strptime(s, f))
        except: continue
    return ERR_VALUE
def excel_TIMEVALUE(s):
    if is_error(s): return s
    s = str(s).strip()
    for f in ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p"):
        try:
            dt = datetime.datetime.strptime(s, f)
            return (dt.hour * 3600 + dt.minute * 60 + dt.second) / 86400.0
        except: continue
    return ERR_VALUE
def excel_WEEKDAY(s, t=1): 
    dt = _get_dt(s); vt = to_num(t)
    if is_error(dt): return dt
    if is_error(vt): return vt
    wd = dt.weekday(); t = int(vt)
    if t == 1: return float((wd + 1) % 7 + 1)
    if t == 2: return float(wd + 1)
    if t == 3: return float(wd)
    return float((wd - (t - 11)) % 7 + 1) if 11 <= t <= 17 else ERR_NUM
def excel_WEEKNUM(s, t=1):
    dt = _get_dt(s); vt = to_num(t);
    if is_error(dt): return dt
    if is_error(vt): return vt
    t = int(vt); 
    if t == 21: return float(dt.isocalendar()[1])
    jan1 = datetime.datetime(dt.year, 1, 1); off = (jan1.weekday() + 1 - {1:1, 2:0, 11:0, 12:6, 13:5, 14:4, 15:3, 16:2, 17:1}.get(t, 1)) % 7
    return float(((dt - jan1).days + off) // 7 + 1)
def excel_DATEDIF(s, e, u):
    ds, de = _get_dt(s), _get_dt(e)
    if is_error(ds): return ds
    if is_error(de): return de
    if ds > de: return ERR_NUM
    u = str(u).upper(); res = 0.0
    if u == "Y": res = de.year - ds.year - ((de.month, de.day) < (ds.month, ds.day))
    elif u == "M": res = (de.year - ds.year) * 12 + de.month - ds.month - (de.day < ds.day)
    elif u == "D": res = (de - ds).days
    return float(res)
def excel_EOMONTH(s, m):
    dt = _get_dt(s); vm = to_num(m)
    if is_error(dt): return dt
    if is_error(vm): return vm
    m = int(vm); tm = dt.month - 1 + m; ny, nm = dt.year + tm // 12, (tm % 12) + 2
    if nm > 12: ny, nm = ny + 1, 1
    return to_excel_date(datetime.datetime(ny, nm, 1) - timedelta(days=1))
def excel_EDATE(s, m):
    dt = _get_dt(s); vm = to_num(m)
    if is_error(dt): return dt
    if is_error(vm): return vm
    m = int(vm); tm = dt.month - 1 + m; ny, nm = dt.year + tm // 12, (tm % 12) + 1
    ld = (datetime.datetime(ny + (nm // 12), (nm % 12) + 1, 1) - timedelta(days=1)).day
    return to_excel_date(datetime.datetime(ny, nm, min(dt.day, ld)))
def excel_TIME(h, m, s):
    vh, vm, vs = to_num(h), to_num(m), to_num(s)
    if is_error(vh): return vh
    if is_error(vm): return vm
    if is_error(vs): return vs
    return ((int(vh)*3600 + int(vm)*60 + int(vs)) % 86400) / 86400.0

# Combinatorics
def excel_FACT(n): v = to_num(n); return v if is_error(v) else (math.factorial(int(v)) if v >= 0 else ERR_NUM)
def excel_FACTDOUBLE(v):
    n = to_num(v); 
    if is_error(n): return n
    n = int(n)
    if n < -1: return ERR_NUM
    if n < 0: return 1.0
    res = 1.0; [res := res * i for i in range(n, 0, -2)]; return float(res)
def excel_COMBIN(n, k):
    vn, vk = to_num(n), to_num(k)
    if is_error(vn) or is_error(vk): return vn if is_error(vn) else vk
    vn, vk = int(vn), int(vk)
    return float(math.comb(vn, vk)) if vn >= vk >= 0 else ERR_NUM
def excel_COMBINA(n, k): vn, vk = int(to_num(n)), int(to_num(k)); return excel_COMBIN(vn + vk - 1, vk)
def excel_PERMUT(n, k):
    vn, vk = to_num(n), to_num(k)
    if is_error(vn) or is_error(vk): return vn if is_error(vn) else vk
    vn, vk = int(vn), int(vk)
    return float(math.perm(vn, vk)) if vn >= vk >= 0 else ERR_NUM
def excel_PERMUTATIONA(n, k): vn, vk = int(to_num(n)), int(to_num(k)); return float(vn ** vk) if vn >= 0 and vk >= 0 else ERR_NUM
def excel_MULTINOMIAL(*args):
    nums = [int(to_num(x)) for x in flat_list(args) if not is_error(to_num(x))]
    if any(x < 0 for x in nums): return ERR_NUM
    res = math.factorial(sum(nums))
    for n in nums: res //= math.factorial(n)
    return float(res)

# Context-Aware Functions
def excel_ROW(*args, context=None, engine=None, current_addr=None):
    ref = args[0] if len(args) > 0 else None
    if ref is None:
        if current_addr:
            _, row = _parse_addr(current_addr)
            return float(row)
        return 1.0
    if isinstance(ref, dict):
        if ref.get('type') == 'cell_reference': return float(_parse_addr(ref['address'])[1])
        if ref.get('type') == 'range_reference':
            r1, r2 = _parse_addr(ref['start']['address'])[1], _parse_addr(ref['end']['address'])[1]
            return [[float(r)] for r in range(min(r1, r2), max(r1, r2) + 1)]
    return ERR_VALUE

def excel_COLUMN(*args, context=None, engine=None, current_addr=None):
    ref = args[0] if len(args) > 0 else None
    if ref is None:
        if current_addr: return float(_col_to_int(_parse_addr(current_addr)[0]))
        return 1.0
    if isinstance(ref, dict):
        if ref.get('type') == 'cell_reference': return float(_col_to_int(_parse_addr(ref['address'])[0]))
        if ref.get('type') == 'range_reference':
            c1, c2 = _col_to_int(_parse_addr(ref['start']['address'])[0]), _col_to_int(_parse_addr(ref['end']['address'])[0])
            return [[float(c) for c in range(min(c1, c2), max(c1, c2) + 1)]]
    return ERR_VALUE

def excel_ROWS(*args, context=None, engine=None, current_addr=None):
    if not args: return ERR_VALUE
    ref = args[0]
    if isinstance(ref, dict):
        if ref.get('type') == 'cell_reference': return 1.0
        if ref.get('type') == 'range_reference': return float(abs(_parse_addr(ref['end']['address'])[1] - _parse_addr(ref['start']['address'])[1]) + 1)
    if isinstance(ref, list): return float(len(ref))
    return 1.0

def excel_COLUMNS(*args, context=None, engine=None, current_addr=None):
    if not args: return ERR_VALUE
    ref = args[0]
    if isinstance(ref, dict):
        if ref.get('type') == 'cell_reference': return 1.0
        if ref.get('type') == 'range_reference': return float(abs(_col_to_int(_parse_addr(ref['end']['address'])[0]) - _col_to_int(_parse_addr(ref['start']['address'])[0])) + 1)
    if isinstance(ref, list): return float(len(ref[0])) if ref and isinstance(ref[0], list) else 1.0
    return 1.0

def excel_OFFSET(*args, context=None, engine=None, current_addr=None):
    if len(args) < 3: return ERR_VALUE
    ref, rows, cols = args[0], int(to_num(args[1])), int(to_num(args[2]))
    height = int(to_num(args[3])) if len(args) > 3 and args[3] is not None else None
    width = int(to_num(args[4])) if len(args) > 4 and args[4] is not None else None
    if not isinstance(ref, dict): return ERR_VALUE
    if ref.get('type') == 'cell_reference':
        base_addr, base_sheet, base_h, base_w = ref['address'], ref.get('sheet'), 1, 1
    elif ref.get('type') == 'range_reference':
        base_addr, base_sheet = ref['start']['address'], ref['start'].get('sheet')
        c1, r1 = _parse_addr(ref['start']['address']); c2, r2 = _parse_addr(ref['end']['address'])
        base_h, base_w = abs(r2 - r1) + 1, abs(_col_to_int(c2) - _col_to_int(c1)) + 1
    else: return ERR_VALUE
    c_str, r = _parse_addr(base_addr); c_int = _col_to_int(c_str)
    target_r, target_c = r + rows, c_int + cols
    if target_r <= 0 or target_c <= 0: return ERR_REF
    h, w = (height if height is not None else base_h), (width if width is not None else base_w)
    if h <= 0 or w <= 0: return ERR_REF
    start_cell = {"type": "cell_reference", "sheet": base_sheet, "address": f"{_int_to_col(target_c)}{target_r}"}
    if h == 1 and w == 1: return engine.evaluate(start_cell, context, current_addr)
    end_cell = {"type": "cell_reference", "sheet": base_sheet, "address": f"{_int_to_col(target_c + w - 1)}{target_r + h - 1}"}
    return engine.evaluate({"type": "range_reference", "start": start_cell, "end": end_cell}, context, current_addr)

def excel_INDIRECT(*args, context=None, engine=None, current_addr=None):
    if not args: return ERR_VALUE
    s = str(args[0])
    try:
        ast = engine.parser.parse(f"={s}")
        return engine.evaluate(ast, context, current_addr)
    except: return ERR_REF

def excel_LOOKUP(*args, context=None, engine=None, current_addr=None):
    if len(args) < 2: return ERR_VALUE
    v, v1, v2 = args[0], args[1], (args[2] if len(args) > 2 else None)
    if v2 is None:
        if not isinstance(v1, list): return ERR_NA
        r, c = len(v1), (len(v1[0]) if len(v1) > 0 and isinstance(v1[0], list) else 1)
        search_area, result_area = (v1[0], v1[-1]) if c > r else ([x[0] for x in v1], [x[-1] for x in v1])
    else:
        search_area, result_area = flat_list([v1]), flat_list([v2])
    best_idx = -1
    for i, x in enumerate(search_area):
        if x == v: best_idx = i
        elif type(x) == type(v) and x < v: best_idx = i
        elif isinstance(x, (int, float)) and isinstance(v, (int, float)) and x < v: best_idx = i
        elif (isinstance(x, (int, float)) and isinstance(v, (int, float)) and x > v) or (type(x) == type(v) and x > v): break
    return result_area[min(best_idx, len(result_area)-1)] if best_idx != -1 else ERR_NA

def excel_VLOOKUP(v, t, i, r=True):
    if not isinstance(t, list) or not t: return ERR_NA
    idx = int(to_num(i)) - 1
    if idx < 0 or idx >= len(t[0]): return ERR_REF
    best_row = -1
    for ri, row in enumerate(t):
        x = row[0]
        if bool(r):
            if x == v: best_row = ri
            elif type(x) == type(v) and x < v: best_row = ri
            elif isinstance(x, (int, float)) and isinstance(v, (int, float)) and x < v: best_row = ri
            elif (type(x) == type(v) and x > v) or (isinstance(x, (int, float)) and isinstance(v, (int, float)) and x > v): break
        else:
            if x == v: return row[idx]
    return t[best_row][idx] if bool(r) and best_row != -1 else ERR_NA

def excel_HLOOKUP(v, t, i, r=True):
    if not isinstance(t, list) or not t: return ERR_NA
    idx = int(to_num(i)) - 1
    if idx < 0 or idx >= len(t): return ERR_REF
    best_col = -1
    for ci, x in enumerate(t[0]):
        if bool(r):
            if x == v: best_col = ci
            elif type(x) == type(v) and x < v: best_col = ci
            elif isinstance(x, (int, float)) and isinstance(v, (int, float)) and x < v: best_col = ci
            elif (type(x) == type(v) and x > v) or (isinstance(x, (int, float)) and isinstance(v, (int, float)) and x > v): break
        else:
            if x == v: return t[idx][ci]
    return t[idx][best_col] if bool(r) and best_col != -1 else ERR_NA

def excel_ADDRESS(r, c, a=1, m=True, s=None, context=None, engine=None, current_addr=None):
    col = _int_to_col(int(to_num(c))); row = int(to_num(r))
    if row <= 0 or int(to_num(c)) <= 0: return ERR_VALUE
    res = f"{'$' if a in (1,2) else ''}{col}{'$' if a in (1,3) else ''}{row}"
    return f"'{s}'!{res}" if s else res

def excel_INDEX(a, r, c=1):
    try:
        r_idx = int(to_num(r)) - 1; c_idx = int(to_num(c)) - 1
        if isinstance(a, list):
            if isinstance(a[0], list): return a[r_idx][c_idx]
            return a[r_idx]
        return a
    except: return ERR_REF

def excel_SUMPRODUCT(*args):
    if not args: return 0.0
    arrs = [flat_list([a]) for a in args]; size = len(arrs[0])
    if any(len(a) != size for a in arrs): return ERR_VALUE
    total = 0.0
    for i in range(size):
        prod = 1.0
        for a in arrs:
            v = a[i]
            if is_error(v): return v
            # SUMPRODUCT treats non-numeric as 0
            try: nv = float(v) if not isinstance(v, (bool, str)) else 0.0
            except: nv = 0.0
            prod *= nv
        total += prod
    return total

def excel_SUMSQ(*args):
    total = 0.0
    for arg in args:
        if isinstance(arg, list):
            for x in flat_list([arg]):
                if is_error(x): return x
                try: total += float(x)**2 if not isinstance(x, (bool, str)) else 0.0
                except: pass
        else:
            if is_error(arg): return arg
            v = to_num(arg)
            if is_error(v): return v
            total += v*v
    return total

def excel_SUMX2MY2(x, y):
    xf, yf = flat_list([x]), flat_list([y])
    if len(xf) != len(yf): return ERR_NA
    total = 0.0
    for vx, vy in zip(xf, yf):
        if is_error(vx) or is_error(vy): return vx if is_error(vx) else vy
        nx, ny = to_num(vx), to_num(vy)
        if isinstance(nx, float) and isinstance(ny, float): total += nx**2 - ny**2
    return total

def excel_SUMX2PY2(x, y):
    xf, yf = flat_list([x]), flat_list([y])
    if len(xf) != len(yf): return ERR_NA
    total = 0.0
    for vx, vy in zip(xf, yf):
        if is_error(vx) or is_error(vy): return vx if is_error(vx) else vy
        nx, ny = to_num(vx), to_num(vy)
        if isinstance(nx, float) and isinstance(ny, float): total += nx**2 + ny**2
    return total

def excel_SUMXMY2(x, y):
    xf, yf = flat_list([x]), flat_list([y])
    if len(xf) != len(yf): return ERR_NA
    total = 0.0
    for vx, vy in zip(xf, yf):
        if is_error(vx) or is_error(vy): return vx if is_error(vx) else vy
        nx, ny = to_num(vx), to_num(vy)
        if isinstance(nx, float) and isinstance(ny, float): total += (nx - ny)**2
    return total

def excel_HYPERLINK(u, f=None): return f if f is not None else u

# Function Registry
NAMES = ["ABS", "ACOS", "ACOSH", "ADDRESS", "AND", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH", "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "BASE", "BIN2DEC", "BIN2HEX", "BIN2OCT", "BITAND", "BITLSHIFT", "BITNOT", "BITOR", "BITRSHIFT", "BITXOR", "CEILING", "CEILING.MATH", "CEILING.PRECISE", "CHAR", "CHOOSE", "CLEAN", "CODE", "COLUMN", "COLUMNS", "COMBIN", "COMBINA", "CONCAT", "CONCATENATE", "COS", "COSH", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "DATE", "DATEDIF", "DATEVALUE", "DAY", "DAYS", "DEC2BIN", "DEC2HEX", "DEC2OCT", "DECIMAL", "DEGREES", "EDATE", "EOMONTH", "EVEN", "EXACT", "EXP", "FACT", "FACTDOUBLE", "FALSE", "FIND", "FLOOR", "FLOOR.MATH", "FLOOR.PRECISE", "FV", "GCD", "HEX2BIN", "HEX2DEC", "HLOOKUP", "HOUR", "IF", "IFERROR", "IFNA", "IFS", "INDEX", "INDIRECT", "INT", "IRR", "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISREF", "ISTEXT", "LARGE", "LCM", "LEFT", "LEN", "LN", "LOG", "LOG10", "LOOKUP", "LOWER", "MATCH", "MAX", "MAXA", "MAXIFS", "MDETERM", "MEDIAN", "MID", "MIN", "MINA", "MINIFS", "MINUTE", "MINVERSE", "MMULT", "MOD", "MODE", "MONTH", "MROUND", "MULTINOMIAL", "N", "NA", "NOT", "NPER", "NPV", "OCT2BIN", "OCT2DEC", "ODD", "OFFSET", "OR", "PERCENTILE", "PERMUT", "PERMUTATIONA", "PI", "PMT", "POWER", "PRODUCT", "PROPER", "PV", "QUOTIENT", "RADIANS", "RANK", "RATE", "REPLACE", "REPT", "RIGHT", "ROUND", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS", "SEARCH", "SECOND", "SIGN", "SIN", "SINH", "SMALL", "SQRT", "STDEV", "STDEVP", "SUBSTITUTE", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "SWITCH", "TAN", "TANH", "TEXT", "TIME", "TIMEVALUE", "TRANSPOSE", "TRIM", "TRUE", "TRUNC", "TYPE", "UPPER", "VALUE", "VAR", "VARP", "VLOOKUP", "WEEKDAY", "WEEKNUM", "XOR", "YEAR"]

FUNCTIONS = {}
for name in NAMES:
    py_name = f"excel_{name.replace('.', '_').upper()}"
    func = globals().get(py_name)
    FUNCTIONS[name] = func if func else (lambda *a, n=name, **kw: ERR_NAME)

FUNCTIONS["TRUE"] = lambda: True; FUNCTIONS["FALSE"] = lambda: False; FUNCTIONS["NA"] = lambda: ERR_NA; FUNCTIONS["PI"] = excel_PI
