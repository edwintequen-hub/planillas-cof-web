import os
import platform


def _leer_entero(ruta):
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            valor = f.read().strip()

        if valor == "max":
            return None

        return int(valor)

    except Exception:
        return None


def _leer_proc_status():
    datos = {}

    try:
        with open(
            "/proc/self/status",
            "r",
            encoding="utf-8"
        ) as f:

            for linea in f:

                if ":" not in linea:
                    continue

                clave, valor = linea.split(":", 1)

                if clave in {
                    "VmRSS",
                    "VmHWM",
                    "VmSize"
                }:
                    datos[clave] = valor.strip()

    except Exception:
        pass

    return datos


def memoria_actual(etapa):
    proc = _leer_proc_status()

    actual = _leer_entero(
        "/sys/fs/cgroup/memory.current"
    )

    limite = _leer_entero(
        "/sys/fs/cgroup/memory.max"
    )

    print(
        "[MEMORIA_DIAG]",
        f"etapa={etapa}",
        f"pid={os.getpid()}",
        f"sistema={platform.system()}",
        f"VmRSS={proc.get('VmRSS', 'N/D')}",
        f"VmHWM={proc.get('VmHWM', 'N/D')}",
        f"VmSize={proc.get('VmSize', 'N/D')}",
        f"cgroup_actual={actual if actual is not None else 'N/D'}",
        f"cgroup_limite={limite if limite is not None else 'N/D'}",
        flush=True
    )
