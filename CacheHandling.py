import pickle

CACHE_FILE = 'cache.pkl'


def save_cache(cache):
    with open(CACHE_FILE, 'wb') as f:
        pickle.dump(cache, f)


def load_cache(default=None):
    try:
        with open(CACHE_FILE, 'rb') as f:
            cache = pickle.load(f)
        if isinstance(cache, dict):
            return cache
    except (FileNotFoundError, pickle.UnpicklingError, EOFError):
        pass
    return default


def save_last_used_font_size(font_size):
    cache = load_cache({})
    cache['font_size'] = font_size
    save_cache(cache)


def load_last_used_font_size():
    cache = load_cache({})
    return str(int(cache.get('font_size', 14)))  # Convert the font size to a string


def save_last_used_directory(directory):
    cache = load_cache({})
    cache['directory'] = directory
    save_cache(cache)


def load_last_used_directory():
    cache = load_cache({})
    return cache.get('directory')


def save_last_used_export_directory(directory):
    cache = load_cache({})
    cache['export_directory'] = directory
    save_cache(cache)


def load_last_used_export_directory():
    cache = load_cache({})
    return cache.get('export_directory')


def save_last_used_theme(theme):
    cache = load_cache({})
    cache['theme'] = theme
    save_cache(cache)


def load_last_used_theme():
    cache = load_cache({})
    return cache.get('theme')
