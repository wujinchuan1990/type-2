def construction_text(instrument, components, usage, custom=""):
    if not instrument.strip():
        raise ValueError("构造和应用场合：当前大纲的计量器具名称缺失或不唯一。")
    if not components or any(not name.strip() for name in components):
        raise ValueError("构造和应用场合：当前大纲的具体零部件名称未完整识别。")
    target = custom.strip() if usage == "其他" else usage
    if usage not in ("用水", "液体", "气体", "燃气", "其他") or not target:
        raise ValueError("构造和应用场合：请选择计量用途；选择“其他”时须填写具体用途。")
    joined = "、".join(components[:-1]) + "和" + components[-1] if len(components) > 1 else components[0]
    return f"{instrument.strip()}主要由{joined}组成，适用于{target}计量。"
