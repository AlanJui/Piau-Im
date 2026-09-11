def to_ms(v):
    if v is None or v == "":
        return 0
    if hasattr(v, "hour"):
        return (
            (v.hour * 3600 + v.minute * 60 + v.second) * 1000
            + int(getattr(v, "microsecond", 0) / 1000)
        )
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    text = str(v).strip()
    if ":" in text:
        h, m, rest = text.split(":")
        sec, ms = rest.replace(".", ",").split(",")
        return (int(h) * 3600 + int(m) * 60 + int(sec)) * 1000 + int(ms)
    return float(text.replace(",", ""))

def ms_to_srt(ms):
    ms = int(round(float(ms)))
    sign = "-" if ms < 0 else ""
    ms = abs(ms)
    h, rem = divmod(ms, 3600000)
    mi, rem = divmod(rem, 60000)
    sec, milli = divmod(rem, 1000)
    return f"{sign}{h:02d}:{mi:02d}:{sec:02d},{milli:03d}"

ms_to_srt(to_ms(xl("B2")) + to_ms(xl("D2")))
