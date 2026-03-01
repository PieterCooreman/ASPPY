from .vb_errors import raise_runtime
from .vb_runtime import vbs_cbool, vbs_cstr
from .vm.values import VBEmpty, VBNull, VBNothing

import datetime as _dt
import math as _math
import os
import struct as _struct
import time as _time

from decimal import Decimal, ROUND_HALF_EVEN

# Re-implement simple math helpers to avoid circular imports or missing definitions
def Sgn(number):
    try:
        x = float(number)
    except:
        x = 0.0
    if x > 0: return 1
    if x < 0: return -1
    return 0

def Sin(number):
    try:
        return _math.sin(float(number))
    except:
        raise_runtime('INVALID_PROC_CALL')

def Tan(number):
    try:
        return _math.tan(float(number))
    except:
        raise_runtime('INVALID_PROC_CALL')

# Stub remaining builtins to satisfy import
def ScriptEngine(): return "VBScript"
def ScriptEngineBuildVersion(): return 1337
def ScriptEngineMajorVersion(): return 5
def ScriptEngineMinorVersion(): return 8
def GetObject(*args): raise_runtime('OBJECT_NOT_SUPPORT', "GetObject is not implemented")

# Date functions
def Weekday(date, firstdayofweek=1):
    from .vb_datetime import Weekday as _W
    return _W(date, firstdayofweek)
def WeekdayName(weekday, abbreviate=False, firstdayofweek=1):
    from .vb_datetime import WeekdayName as _W
    return _W(weekday, abbreviate, firstdayofweek)
def MonthName(month, abbreviate=False):
    from .vb_datetime import MonthName as _M
    return _M(month, abbreviate)
def Day(date):
    from .vb_datetime import Day as _D
    return _D(date)
def Month(date):
    from .vb_datetime import Month as _M
    return _M(date)
def Year(date):
    from .vb_datetime import Year as _Y
    return _Y(date)
def Hour(time):
    from .vb_datetime import Hour as _H
    return _H(time)
def Minute(time):
    from .vb_datetime import Minute as _M
    return _M(time)
def Second(time):
    from .vb_datetime import Second as _S
    return _S(time)
def DateAdd(interval, number, date):
    from .vb_datetime import DateAdd as _D
    return _D(interval, number, date)
def DateDiff(interval, date1, date2, firstdayofweek=1, firstweekofyear=1):
    from .vb_datetime import DateDiff as _D
    return _D(interval, date1, date2, firstdayofweek, firstweekofyear)
def DatePart(interval, date, firstdayofweek=1, firstweekofyear=1):
    from .vb_datetime import DatePart as _D
    return _D(interval, date, firstdayofweek, firstweekofyear)
def DateSerial(year, month, day):
    from .vb_datetime import DateSerial as _D
    return _D(year, month, day)
def DateValue(date):
    from .vb_datetime import DateValue as _D
    return _D(date)
def TimeSerial(hour, minute, second):
    from .vb_datetime import TimeSerial as _T
    return _T(hour, minute, second)
def TimeValue(time):
    from .vb_datetime import TimeValue as _T
    return _T(time)
def Now():
    from .vb_datetime import Now as _N
    return _N()
def Date():
    from .vb_datetime import Date as _D
    return _D()
def Time():
    from .vb_datetime import Time as _T
    return _T()
def Timer():
    from .vb_datetime import Timer as _T
    return _T()

# VBScript-like RNG state (24-bit LCG) to mirror VM behavior.
_RND_MOD = 16777216
_RND_A = 1140671485
_RND_C = 12820163
_rnd_state = int.from_bytes(os.urandom(4), 'little') % _RND_MOD
_rnd_last = _rnd_state / float(_RND_MOD)

def Randomize(seed=None):
    global _rnd_state, _rnd_last
    if seed is None or seed == "":
        try:
            s = int(float(Timer()) * 1000000.0)
        except Exception:
            s = int(_time.time() * 1000000.0)
    else:
        try:
            s = int(float(seed) * 1000000)
        except Exception:
            s = int(_time.time() * 1000000.0)
    _rnd_state = int(s) % _RND_MOD
    _rnd_last = _rnd_state / float(_RND_MOD)

def Rnd(number=None):
    global _rnd_state, _rnd_last
    n = None
    if number is not None:
        try:
            n = float(number)
        except Exception:
            n = 0.0

    if n is not None and n < 0:
        b = _struct.pack('<f', float(n))
        bits = _struct.unpack('<I', b)[0]
        seed24 = (bits & 0xFFFFFF) | ((bits >> 24) & 0xFF)
        _rnd_state = int(seed24) % _RND_MOD
        _rnd_state = (_rnd_state * _RND_A + _RND_C) % _RND_MOD
        _rnd_last = _rnd_state / float(_RND_MOD)
        return _rnd_last

    if n is not None and n == 0:
        return _rnd_last

    _rnd_state = (_rnd_state * _RND_A + _RND_C) % _RND_MOD
    _rnd_last = _rnd_state / float(_RND_MOD)
    return _rnd_last

def FormatDateTime(date, namedformat=0):
    from .vb_datetime import FormatDateTime as _F
    return _F(date, namedformat)
