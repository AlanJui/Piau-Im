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

old_show = xl("B5")
dur = xl("B6")
new_show = xl("B9")
new_dur = xl("B10")
offset = srt_to_ms(new_show) - srt_to_ms(old_show)

pd.DataFrame({
    "項目": ["目前Hide", "新Hide", "位移毫秒", "位移SRT"],
    "值": [
        ms_to_srt(srt_to_ms(old_show) + dur),
        ms_to_srt(srt_to_ms(new_show) + new_dur),
        offset,
        ms_to_srt(offset),
    ]
})
