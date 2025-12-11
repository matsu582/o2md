# -*- coding: utf-8 -*-
"""
Office Math Markup Language (OMML) から LaTeX への変換

OMML 要素を解析し、対応する LaTeX 形式に変換します。
markitdown (MIT License) の実装を参考にしています。
"""

from xml.etree import ElementTree as ET

from .latex_dict import (
    CHARS,
    CHR,
    CHR_BO,
    CHR_DEFAULT,
    POS,
    POS_DEFAULT,
    SUB,
    SUP,
    F,
    F_DEFAULT,
    T,
    FUNC,
    D,
    D_DEFAULT,
    RAD,
    RAD_DEFAULT,
    ARR,
    LIM_FUNC,
    LIM_TO,
    LIM_UPP,
    M,
    BRK,
    BLANK,
    BACKSLASH,
    ALN,
    FUNC_PLACE,
)

# OMML 名前空間
OMML_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def load(stream):
    """ストリームから OMML を読み込み、LaTeX に変換"""
    tree = ET.parse(stream)
    for omath in tree.findall(OMML_NS + "oMath"):
        yield oMath2Latex(omath)


def load_string(string):
    """文字列から OMML を読み込み、LaTeX に変換"""
    root = ET.fromstring(string)
    for omath in root.findall(OMML_NS + "oMath"):
        yield oMath2Latex(omath)


def escape_latex(strs):
    """LaTeX の特殊文字をエスケープ"""
    last = None
    new_chr = []
    strs = strs.replace(r"\\", "\\")
    for c in strs:
        if (c in CHARS) and (last != BACKSLASH):
            new_chr.append(BACKSLASH + c)
        else:
            new_chr.append(c)
        last = c
    return BLANK.join(new_chr)


def get_val(key, default=None, store=CHR):
    """辞書から値を取得、キーがない場合はデフォルト値を返す"""
    if key is not None:
        return key if not store else store.get(key, key)
    else:
        return default


class Tag2Method:
    """タグ名からメソッドを呼び出す基底クラス"""
    
    def call_method(self, elm, stag=None):
        """タグに対応するメソッドを呼び出す"""
        getmethod = self.tag2meth.get
        if stag is None:
            stag = elm.tag.replace(OMML_NS, "")
        method = getmethod(stag)
        if method:
            return method(self, elm)
        else:
            return None

    def process_children_list(self, elm, include=None):
        """子要素を処理し、リストとして返す"""
        for _e in list(elm):
            if OMML_NS not in _e.tag:
                continue
            stag = _e.tag.replace(OMML_NS, "")
            if include and (stag not in include):
                continue
            t = self.call_method(_e, stag=stag)
            if t is None:
                t = self.process_unknow(_e, stag)
                if t is None:
                    continue
            yield (stag, t, _e)

    def process_children_dict(self, elm, include=None):
        """子要素を処理し、辞書として返す"""
        latex_chars = dict()
        for stag, t, e in self.process_children_list(elm, include):
            latex_chars[stag] = t
        return latex_chars

    def process_children(self, elm, include=None):
        """子要素を処理し、文字列として返す"""
        return BLANK.join(
            (
                t if not isinstance(t, Tag2Method) else str(t)
                for stag, t, e in self.process_children_list(elm, include)
            )
        )

    def process_unknow(self, elm, stag):
        """未知のタグを処理"""
        return None


class Pr(Tag2Method):
    """プロパティ要素を処理するクラス"""
    text = ""

    __val_tags = ("chr", "pos", "begChr", "endChr", "type")
    __innerdict = None

    def __init__(self, elm):
        self.__innerdict = {}
        self.text = self.process_children(elm)

    def __str__(self):
        return self.text

    def __unicode__(self):
        return self.__str__(self)

    def __getattr__(self, name):
        return self.__innerdict.get(name, None)

    def do_brk(self, elm):
        """改行を処理"""
        self.__innerdict["brk"] = BRK
        return BRK

    def do_common(self, elm):
        """共通プロパティを処理"""
        stag = elm.tag.replace(OMML_NS, "")
        if stag in self.__val_tags:
            t = elm.get("{0}val".format(OMML_NS))
            self.__innerdict[stag] = t
        return None

    tag2meth = {
        "brk": do_brk,
        "chr": do_common,
        "pos": do_common,
        "begChr": do_common,
        "endChr": do_common,
        "type": do_common,
    }


class oMath2Latex(Tag2Method):
    """oMath 要素を LaTeX に変換するクラス"""

    _t_dict = T
    __direct_tags = ("box", "sSub", "sSup", "sSubSup", "num", "den", "deg", "e")

    def __init__(self, element):
        self._latex = self.process_children(element)

    def __str__(self):
        return self.latex

    def __unicode__(self):
        return self.__str__(self)

    def process_unknow(self, elm, stag):
        """未知のタグを処理"""
        if stag in self.__direct_tags:
            return self.process_children(elm)
        elif stag[-2:] == "Pr":
            return Pr(elm)
        else:
            return None

    @property
    def latex(self):
        """LaTeX 文字列を返す"""
        return self._latex

    def do_acc(self, elm):
        """アクセント関数を処理"""
        c_dict = self.process_children_dict(elm)
        latex_s = get_val(
            c_dict["accPr"].chr, default=CHR_DEFAULT.get("ACC_VAL"), store=CHR
        )
        return latex_s.format(c_dict["e"])

    def do_bar(self, elm):
        """バー関数を処理"""
        c_dict = self.process_children_dict(elm)
        pr = c_dict["barPr"]
        latex_s = get_val(pr.pos, default=POS_DEFAULT.get("BAR_VAL"), store=POS)
        return pr.text + latex_s.format(c_dict["e"])

    def do_d(self, elm):
        """区切り記号オブジェクトを処理"""
        c_dict = self.process_children_dict(elm)
        pr = c_dict["dPr"]
        null = D_DEFAULT.get("null")
        s_val = get_val(pr.begChr, default=D_DEFAULT.get("left"), store=T)
        e_val = get_val(pr.endChr, default=D_DEFAULT.get("right"), store=T)
        return pr.text + D.format(
            left=null if not s_val else escape_latex(s_val),
            text=c_dict["e"],
            right=null if not e_val else escape_latex(e_val),
        )

    def do_spre(self, elm):
        """前置上付き・下付きオブジェクトを処理（未サポート）"""
        pass

    def do_sub(self, elm):
        """下付き文字を処理"""
        text = self.process_children(elm)
        return SUB.format(text)

    def do_sup(self, elm):
        """上付き文字を処理"""
        text = self.process_children(elm)
        return SUP.format(text)

    def do_f(self, elm):
        """分数オブジェクトを処理"""
        c_dict = self.process_children_dict(elm)
        pr = c_dict.get("fPr")
        if pr:
            latex_s = get_val(pr.type, default=F_DEFAULT, store=F)
            prefix = pr.text
        else:
            # fPr が存在しない場合はデフォルトの分数形式を使用
            latex_s = F_DEFAULT
            prefix = ""
        return prefix + latex_s.format(num=c_dict.get("num"), den=c_dict.get("den"))

    def do_func(self, elm):
        """関数適用オブジェクトを処理 (sin, cos など)"""
        c_dict = self.process_children_dict(elm)
        func_name = c_dict.get("fName")
        return func_name.replace(FUNC_PLACE, c_dict.get("e"))

    def do_fname(self, elm):
        """関数名を処理"""
        latex_chars = []
        for stag, t, e in self.process_children_list(elm):
            if stag == "r":
                if FUNC.get(t):
                    latex_chars.append(FUNC[t])
                else:
                    # 未知の関数名はそのまま使用
                    latex_chars.append(f"\\mathrm{{{t}}}")
            else:
                latex_chars.append(t)
        t = BLANK.join(latex_chars)
        return t if FUNC_PLACE in t else t + FUNC_PLACE

    def do_groupchr(self, elm):
        """グループ文字オブジェクトを処理"""
        c_dict = self.process_children_dict(elm)
        pr = c_dict["groupChrPr"]
        latex_s = get_val(pr.chr)
        return pr.text + latex_s.format(c_dict["e"])

    def do_rad(self, elm):
        """根号オブジェクトを処理"""
        c_dict = self.process_children_dict(elm)
        text = c_dict.get("e")
        deg_text = c_dict.get("deg")
        if deg_text:
            return RAD.format(deg=deg_text, text=text)
        else:
            return RAD_DEFAULT.format(text=text)

    def do_eqarr(self, elm):
        """配列オブジェクトを処理"""
        return ARR.format(
            text=BRK.join(
                [t for stag, t, e in self.process_children_list(elm, include=("e",))]
            )
        )

    def do_limlow(self, elm):
        """下極限オブジェクトを処理"""
        t_dict = self.process_children_dict(elm, include=("e", "lim"))
        latex_s = LIM_FUNC.get(t_dict["e"])
        if not latex_s:
            # 未知の極限関数はデフォルトの lim を使用
            return f"\\lim_{{{t_dict.get('lim')}}}"
        else:
            return latex_s.format(lim=t_dict.get("lim"))

    def do_limupp(self, elm):
        """上極限オブジェクトを処理"""
        t_dict = self.process_children_dict(elm, include=("e", "lim"))
        return LIM_UPP.format(lim=t_dict.get("lim"), text=t_dict.get("e"))

    def do_lim(self, elm):
        """極限の下限/上限を処理"""
        return self.process_children(elm).replace(LIM_TO[0], LIM_TO[1])

    def do_m(self, elm):
        """行列オブジェクトを処理"""
        rows = []
        for stag, t, e in self.process_children_list(elm):
            if stag == "mPr":
                pass
            elif stag == "mr":
                rows.append(t)
        return M.format(text=BRK.join(rows))

    def do_mr(self, elm):
        """行列の行を処理"""
        return ALN.join(
            [t for stag, t, e in self.process_children_list(elm, include=("e",))]
        )

    def do_nary(self, elm):
        """n項演算子オブジェクトを処理"""
        res = []
        bo = ""
        for stag, t, e in self.process_children_list(elm):
            if stag == "naryPr":
                bo = get_val(t.chr, store=CHR_BO)
            else:
                res.append(t)
        return bo + BLANK.join(res)

    def do_r(self, elm):
        """テキスト要素を処理"""
        _str = []
        for s in elm.findtext("./{0}t".format(OMML_NS)):
            _str.append(self._t_dict.get(s, s))
        return escape_latex(BLANK.join(_str))

    tag2meth = {
        "acc": do_acc,
        "r": do_r,
        "bar": do_bar,
        "sub": do_sub,
        "sup": do_sup,
        "f": do_f,
        "func": do_func,
        "fName": do_fname,
        "groupChr": do_groupchr,
        "d": do_d,
        "rad": do_rad,
        "eqArr": do_eqarr,
        "limLow": do_limlow,
        "limUpp": do_limupp,
        "lim": do_lim,
        "m": do_m,
        "mr": do_mr,
        "nary": do_nary,
    }
