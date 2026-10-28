"""单位换算：docx 用 EMU/twips，教师看得懂的是毫米和厘米。"""

EMU_PER_MM = 36000
EMU_PER_CM = 360000
TWIPS_PER_MM = 56.69291338582677
PT_PER_MM = 2.8346456692913385


def emu_to_mm(value):
    return None if value is None else round(value / EMU_PER_MM, 1)


def twips_to_mm(value):
    return None if value is None else round(value / TWIPS_PER_MM, 1)


def pt_to_mm(value):
    return None if value is None else round(value / PT_PER_MM, 1)


def mm_to_emu(value):
    return int(round(value * EMU_PER_MM))
