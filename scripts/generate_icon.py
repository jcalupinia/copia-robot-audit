from pathlib import Path


def main() -> int:
    base_dir = Path(__file__).resolve().parent.parent
    png_candidates = [
        base_dir / "LogoAUDIT.png",
        base_dir / "AUDIT_IA_sin_fondo_transparente_FINAL.png",
    ]
    png_path = next((p for p in png_candidates if p.exists()), None)
    if png_path is None:
        print("[WARN] No se encontro PNG para generar icono.")
        return 0

    ico_path = base_dir / "LogoAUDIT.ico"
    try:
        from PIL import Image  # type: ignore
    except Exception:
        print("[WARN] Pillow no disponible; no se genero LogoAUDIT.ico.")
        return 0

    try:
        img = Image.open(png_path).convert("RGBA")
        sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        img.save(ico_path, format="ICO", sizes=sizes)
        print(f"[OK] Icono generado: {ico_path}")
    except Exception as err:
        print(f"[WARN] No se pudo generar icono: {err}")
        return 0

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
