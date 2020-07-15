"""
======================possible encodings======================================================
website: https://www.sljfaq.org/afaq/encodings.html
python doc: https://docs.python.org/2.4/lib/standard-encodings.html

#<~ JIS (Japanese Industrial Standard) character set:
demerit ---> This is mainly a problem with JIS, since its revision process has been somewhat chaotic.
scope ---> Shift JIS is the Microsoft encoding of JIS, standard on Windows and Mac systems. Almost all Japanese web pages used to be encoded in Shift JIS.
       EUC (EUC-JP) is the Unix encoding of JIS. It used to be standard on Linux/BSD systems, and is sometimes used on web pages.
       ISO-2022-JP is the 7-bit JIS encoding, that e-mails are usually sent in. It is rarely used in other contexts.

#<~ Unicode
merit ---> Unicode follows a policy that each new revision must be a strict super-set of previous ones, so version conflicts rarely cause problems.
scope ---> UTF-8 is the Unicode encoding standard in Unix and on the internet
        and UTF-16 the Unicode encoding standard in Windows
        UTF-32 is mostly used for internal representation inside programs, and not for interchange.
details --->
JIS X 0201 - Roman characters (a, b, c ...) and half-width katakana only. Standard prefix: JG.
JIS X 0208 - Core character set. 6355 kanji from level 1 and 2, 524 kana/punctuation/etc. Standard prefix: JH.
JIS X 0212 - Supplemental characters. 5801 rare kanji, 266 non-English European characters. Standard prefix: JJ.
JIS X 0213 - New unified character set. Sometimes called JIS2000, introduced in the year 2000, and meant to replace the 0208/0212 combination. It has no standard prefix.

The above four JIS standards are the important ones, but for reference, here is the meaning of all the JIS X 200 codes:
    JIS X 0201 - Roman/katakana (JG)
    JIS X 0202 - ISO-2022-JP
    JIS X 0203, 0204, 0205, 0206, 0207 - Obsolete/withdrawn standards
    JIS X 0208 - Main kanji character set (JH)
    JIS X 0209 - How to write JIS X 0211 characters
    JIS X 0210 - How to encode numbers
    JIS X 0211 - Standard ASCII control codes
    JIS X 0212 - Supplemental character set (JJ)
    JIS X 0213 - New unified JIS character set
    JIS X 0218 - Definition of standard prefixes (JG, JH, JJ, JK, JL)
    JIS X 0221 - Unicode (JK for UCS-2, JL for UCS-4)

cp932
--->CP932 (code page 932) is an extension of Shift JIS from Microsoft. It adds the NEC/IBM extended characters. This extension is:
NEC special characters (83 characters in SJIS row 13),
NEC-selected IBM extended characters (374 characters in SJIS rows 89..92),
and IBM extended characters (388 characters in SJIS rows 115..119).


euc_jp, euc_jis_2004, euc_jisx0213
--->EUC (Extended Unix Code), also known as UJIS, is an encoding which encodes all the characters of JIS X 0201, JIS X 0208 and JIS X 0212. It is compatible with ASCII, but not with JIS X 0201: EUC does not support 1-byte half-width katakana/punctuation (though it does support it in 2 bytes).

iso2022_jp, iso2022_jp_1, iso2022_jp_2, iso2022_jp_2004, iso2022_jp_3, iso2022_jp_ext
--->The most widely supported encoding for e-mail is the 7-bit ISO-2022-JP standard, which has been used for Japanese e-mail since the beginning. This encoding is almost certain to be understood by a Japanese recipient. It has also been standardized by the JSA under the name "JIS X 0202".


shift-jis
--->Shift JIS is an encoding of the JIS standard which was the standard encoding for Japanese on Microsoft and Apple computers before the advent of Unicode. The selling point of Shift JIS (a.k.a. SJIS) is that, unlike EUC, it is backwards-compatible with not only ASCII, but also JIS X 0201, so Shift JIS can be used to encode both JIS X 0201 and JIS X 0208 (but not JIS X 0212). One-byte half-width katakana/punctuation is valid Shift JIS. Unfortunately, this compatibility means that Shift JIS is the messiest encoding of all.

so, when trying to convert VBA modules origated from SONY JP, the best decoding methods in windows are:
shift_jis_2004 or shift_jisx0213 > shift_jis >  cp932

=============================================================================================
"""

import os, glob

def vba_jp_decoding(file_path, encoding, save2fd_path):
    fn = os.path.split(file_path)[-1]
    save2file_path = os.path.join(save2fd_path, fn)
    with open(file_path, 'r', encoding=encoding) as srcf, open(save2file_path, 'w', encoding='utf-8') as newf:
        for line in srcf.read():
            newf.write(line)
    return
    
def main():
    chrset = ['shift_jis_2004', 'shift_jisx0213', 'shift_jis', 'cp932']
    base_dir = os.path.dirname(__file__)
    src_fd_path = os.path.abspath(os.path.join(base_dir, "vba_mod_src"))
    save2fd_path = os.path.join(base_dir, "converted2fd")
    
    if not os.path.isdir(save2fd_path): os.makedirs(save2fd_path, exist_ok=True)
    src_files = sorted(glob.glob(os.path.join(src_fd_path, '*.bas')))
    encoding = chrset[-1]
    for f in src_files:
        vba_jp_decoding(f, encoding, save2fd_path)

if __name__ == '__main__':
    main()
