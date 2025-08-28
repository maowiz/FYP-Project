from threading import Event


_speaking_event = Event()


def set_speaking(is_speaking: bool) -> None:
    if is_speaking:
        _speaking_event.set()
    else:
        _speaking_event.clear()


def is_speaking() -> bool:
    return _speaking_event.is_set()


