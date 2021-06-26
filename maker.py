import os

from argparse import ArgumentParser

import transposer
import make_docx


def read_data_txt(src):
    ctlist = []
    for l in open(src):

        l = l.strip()
        if len(l) == 0:
            ctlist.append([])
            continue

        e = l.split(',')

        ctline = []
        for p in e:
            if ':' in p:
                try:
                    c, t = p.split(':')
                except ValueError as e:
                    print(e)
                    print(p)
                    exit(1)
                ctline.append([c, t])
            else:
                ctline.append(['', p])

        ctlist.append(ctline)

    return ctlist


if __name__ == "__main__":
    parser = ArgumentParser()
    parser.add_argument('src', help='コード、歌詞データファイル')
    parser.add_argument('-o', '--out', help='出力ファイル名')
    parser.add_argument('-t', '--transpose', help='移調', default=0)
    parser.add_argument('-s', '--song', help='曲名', default='song')
    parser.add_argument('-a', '--artist', help='アーティスト名', default='artist')
    parser.add_argument('--flat', help='移調の時に♭を使う', action='store_true')

    args = parser.parse_args()

    src = args.src
    if args.out is None:
        out = os.path.splitext(os.path.basename(src))[0] + '_out'
    else:
        out = args.out
    transpose = int(args.transpose)

    song = args.song
    artist = args.artist

    read_data = read_data_txt(src)

    use_sharp = not args.flat
    ctlist = transposer.transpose_list(read_data, transpose, use_sharp=use_sharp)

    if '.docx' not in out:
        out += '.docx'

    make_docx.make(out, ctlist, song, artist, write_chord=True)
