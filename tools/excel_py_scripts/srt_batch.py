def srt_to_ms(s):
    text = str(s).strip()
    h, m, rest = text.split(":")
    sec, ms = rest.replace(".", ",").split(",")
    return (int(h) * 3600 + int(m) * 60 + int(sec)) * 1000 + int(ms)

def ms_to_srt(ms):
    ms = int(round(float(ms)))
    sign = "-" if ms < 0 else ""
    ms = abs(ms)
    h, rem = divmod(ms, 3600000)
    mi, rem = divmod(rem, 60000)
    sec, milli = divmod(rem, 1000)
    return f"{sign}{h:02d}:{mi:02d}:{sec:02d},{milli:03d}"

df = xl("A1:F62", headers=True)
old_show = xl("單句!B5")
new_show = xl("單句!B9")
start_no = xl("單句!B14")
offset = srt_to_ms(new_show) - srt_to_ms(old_show)

def shift_one(t, no):
    ms = srt_to_ms(t)
    if no >= start_no:
        ms = ms + offset
    return ms_to_srt(ms)

pd.DataFrame({
    "新Show": [shift_one(s, n) for s, n in zip(df["Show"], df["編號"])],
    "新Hide": [shift_one(s, n) for s, n in zip(df["Hide"], df["編號"])],
})
