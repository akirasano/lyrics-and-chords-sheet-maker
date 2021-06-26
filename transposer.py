import copy

sharp_data = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
flat_data = ['C', 'D♭', 'D', 'E♭', 'E', 'F', 'G♭', 'G', 'A♭', 'A', 'B♭', 'B']


def search_base(chord):
    if '#' in chord:
        base = chord.split('#')[0]
    elif '♭' in chord:
        base = chord.split('♭')[0]
    else:
        base = None
        for q in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            if q in chord:
                base = q
                break
        if base is None:
            raise Exception(f'{chord} is not found in chord list.')

    return base


def shift(base, nshift=0, use_sharp=True):
    if nshift == 0:
        return base

    if '#' in base:
        ca = sharp_data
    elif '♭' in base:
        ca = flat_data
    else:
        if use_sharp:
            ca = sharp_data
        else:
            ca = flat_data

    bi = ca.index(base)
    tbi = (bi + nshift) % 12
    return ca[tbi]


def transpose(chord, *args, **kwargs):
    if chord == '':
        return chord

    base = search_base(chord)
    suffix = chord.replace(base, '')
    transposed_base = shift(base, *args, **kwargs)
    if '/' in chord:
        suffix, bassnote = suffix.split('/')
        transposed_bassnote = shift(bassnote, *args, **kwargs)
        transposed = f'{transposed_base}{suffix}/{transposed_bassnote}'
    else:
        transposed = f'{transposed_base}{suffix}'
    return transposed


def transpose_list(chord_list, nshift, use_sharp=None):
    if nshift == 0:
        return chord_list

    def has_sharp(chord_list):
        for line in chord_list:
            for chord_elem in line:
                chord, lyrics = chord_elem
                if '#' in chord:
                    return True
                elif '♭' in chord:
                    return False
        return True

    transposed_chord_list = copy.deepcopy(chord_list)
    if use_sharp is None:
        use_sharp = has_sharp(chord_list)
    for line in transposed_chord_list:
        for chord_elem in line:
            chord, lyrics = chord_elem
            tc = transpose(
                chord, nshift=nshift, use_sharp=use_sharp)
            chord_elem[0] = tc

    return transposed_chord_list
