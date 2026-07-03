# -*- coding: utf-8 -*-
"""Generate rapport.tex (XeLaTeX/polyglossia) from the student's .docx.
Faithful conversion + new general introduction + rebuilt Chapter 4."""
import zipfile, xml.etree.ElementTree as ET

DOCX = "/root/.claude/uploads/d37b352c-c516-52d1-98da-8d676b1ba3f6/50dfb9dc-_____.docx"
OUT  = "/home/user/MP3Translator/rapport/rapport.tex"
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# ----------------------------------------------------------------------
# 1. Parse docx body into ordered blocks
# ----------------------------------------------------------------------
def para_text(p):
    return ''.join(t.text or '' for t in p.iter(W+'t')).strip()

def para_style(p):
    ppr = p.find(W+'pPr')
    if ppr is None: return ''
    ps = ppr.find(W+'pStyle')
    return ps.get(W+'val','') if ps is not None else ''

def dedup(t):
    # cover/heading runs are often duplicated (textbox + body)
    n=len(t)
    if n>8 and n%2==0 and t[:n//2]==t[n//2:]:
        return t[:n//2]
    return t

root = ET.fromstring(zipfile.ZipFile(DOCX).read('word/document.xml'))
body = root.find(W+'body')
blocks=[]
for el in body:
    tag=el.tag.replace(W,'')
    if tag=='p':
        txt=dedup(para_text(el)); st=para_style(el)
        if txt: blocks.append(('P',st,txt))
    elif tag=='tbl':
        rows=[]
        for tr in el.findall(W+'tr'):
            cells=[dedup(' '.join(para_text(p) for p in tc.findall(W+'p')).strip())
                   for tc in tr.findall(W+'tc')]
            rows.append(cells)
        blocks.append(('T',None,rows))

# ----------------------------------------------------------------------
# 2. LaTeX escaping
# ----------------------------------------------------------------------
def esc(s):
    s=s.replace('\\',r'\textbackslash{}')
    for a,b in [('&',r'\&'),('%',r'\%'),('$',r'\$'),('#',r'\#'),
                ('_',r'\_'),('{',r'\{'),('}',r'\}'),('~',r'\textasciitilde{}'),
                ('^',r'\textasciicircum{}')]:
        s=s.replace(a,b)
    s=s.replace('✓', r'$\checkmark$').replace('�merged','')
    for ch in ['🔴','🟡','🟢','🔵']:
        s=s.replace(ch,'')
    s=s.replace('—','--').replace('–','-')
    import re as _r
    s=_r.sub(r'\*\*(.+?)\*\*', r'\\textbf{\1}', s)   # markdown bold -> \textbf
    return s

# ----------------------------------------------------------------------
# 3. Table emitter (adaptive column specs, RTL-safe, width-fitting)
# ----------------------------------------------------------------------
def emit_table(rows, caption):
    rows=[r for r in rows if any(c.strip() for c in r)]
    if not rows: return ''
    ncols=max(len(r) for r in rows)
    rows=[r+['']*(ncols-len(r)) for r in rows]
    if ncols>=6:
        colspec='l'+' c'*(ncols-1); wide=True; size=r'\scriptsize'
    elif ncols in (4,5):
        colspec='l'+' c'*(ncols-1); wide=True; size=r'\small'
    elif ncols==3:
        colspec='p{6cm} p{4cm} p{4cm}'; wide=False; size=r'\small'
    else:
        colspec='p{3.8cm} p{9.6cm}'; wide=False; size=r'\small'
    out=[]
    out.append(r'\begin{table}[H]')
    out.append(r'\centering')
    if caption: out.append(r'\caption{%s}' % esc(caption))
    out.append(size)
    out.append(r'\renewcommand{\arraystretch}{1.3}')
    out.append(r'\setlength{\tabcolsep}{4pt}')
    inner=[]
    inner.append(r'\begin{tabular}{%s}' % colspec)
    inner.append(r'\toprule')
    hdr=rows[0]
    inner.append(' & '.join(r'\textbf{%s}'%esc(c) for c in hdr)+r' \\')
    inner.append(r'\midrule')
    for r in rows[1:]:
        inner.append(' & '.join(esc(c) for c in r)+r' \\')
    inner.append(r'\bottomrule')
    inner.append(r'\end{tabular}')
    tab='\n'.join(inner)
    if wide:
        out.append(r'\resizebox{\textwidth}{!}{%')
        out.append(tab)
        out.append(r'}')
    else:
        out.append(tab)
    out.append(r'\end{table}')
    return '\n'.join(out)

# ----------------------------------------------------------------------
# 4. Heading heuristics
# ----------------------------------------------------------------------
SEC_PREF=('أولاً','أولًا','ثانياً','ثانيًا','ثالثاً','ثالثًا','رابعاً','رابعًا',
          'خامساً','خامسًا','سادساً','سادسًا','سابعاً','سابعًا','ثامناً','ثامنًا',
          'تاسعاً','تاسعًا','عاشراً')
BED=('البعد الأول','البعد الثاني','البعد الثالث','البعد الرابع','البعد الخامس')
MEHWAR=('المحور الأول','المحور الثاني','المحور الثالث')
SUBQ=('السؤال الأول','السؤال الثاني','السؤال الافتتاحي','السؤال الختامي',
      'سؤال افتتاحي','سؤال ختامي','السؤال الافتتاحي العام')
def clean_lead(t):
    return t.lstrip('.-–—0123456789 ').strip()
def heading_level(t):
    import re as _r
    tt=t.strip()
    if len(tt)<70 and tt.startswith(SEC_PREF): return ('sec', tt)
    if len(tt)<70 and any(tt.startswith(b) for b in BED): return ('sec', tt)
    if len(tt)<85 and any(tt.startswith(m) for m in MEHWAR): return ('sec', tt)
    if tt.startswith('الملحق'): return ('sec', tt)
    if len(tt)<70 and any(tt.startswith(q) for q in SUBQ): return ('sub', tt)
    if tt in ('تمهيد','تمهيد الفصل'): return ('secstar','تمهيد')
    if tt in ('الأسئلة الفرعية','السؤال المركزي'): return ('sub', tt)
    if 'التعريف الإجرائي' in tt and len(tt)<45: return ('sub', tt)
    if tt.startswith('مفهوم الإدراك'): return ('sub', tt)
    if tt.startswith('. ') and len(tt)<80: return ('sub', clean_lead(tt))
    if _r.match(r'^\d+[.\-]\s', tt) and len(tt)<45: return ('sub', clean_lead(tt))
    return (None,None)

RUNIN=('الجواب:','السؤال:','المؤشر المستخلص:','المؤشر المستهدف:','البعد المرتبط:',
       'السند النظري:','الأسئلة:','وظيفته:','الملاحظة التحليلية','الاتجاه العام',
       'الحالات الخاصة','القراءة التحليلية','خلاصة الفصل','ملاحظة منهجية','ملاحظة:',
       'الكلمات المفتاحية:')

# ----------------------------------------------------------------------
# 5. Static LaTeX pieces (hand-written)
# ----------------------------------------------------------------------
PREAMBLE=r"""% =====================================================================
%  بحث لنيل شهادة الإجازة في علم النفس الإكلينيكي
%  إدراك أسلوب المعاملة الوالدية المتسلط وعلاقته بتقدير الذات لدى المراهقين
%
%  المحرّك: XeLaTeX   (على Overleaf: Menu > Settings > Compiler > XeLaTeX)
%  الخطّ العربي: Amiri (متوفّر على Overleaf) ، وإلا FreeSerif.
% =====================================================================
\documentclass[12pt,oneside]{report}
\usepackage{fontspec}
\usepackage{geometry}
\geometry{a4paper, top=2.5cm, bottom=2.5cm, left=2.5cm, right=2.5cm}
\usepackage{setspace}\onehalfspacing
\usepackage{xcolor}
\definecolor{titleblue}{RGB}{7,15,185}
\definecolor{lightgray}{RGB}{238,238,238}
\usepackage{graphicx}
\usepackage{array}
\usepackage{booktabs}
\usepackage{colortbl}
\usepackage{longtable}
\usepackage{amssymb}
\usepackage{float}
\usepackage{enumitem}
\usepackage{hyperref}
\hypersetup{colorlinks=true, linkcolor=titleblue, urlcolor=titleblue,
            unicode=true, bookmarksnumbered=true, pdfborder={0 0 0}}
\usepackage{titlesec}
\titleformat{\chapter}[display]
  {\normalfont\huge\bfseries}
  {\chaptertitlename\ \thechapter}{12pt}{\Huge}
\titlespacing*{\chapter}{0pt}{6pt}{26pt}
\renewcommand{\contentsname}{فهرس المحتويات}
\renewcommand{\listtablename}{لائحة الجداول}
\renewcommand{\listfigurename}{لائحة الأشكال}
\renewcommand{\chaptername}{الفصل}
\renewcommand{\tablename}{جدول}
\renewcommand{\figurename}{شكل}
\setcounter{tocdepth}{2}\setcounter{secnumdepth}{2}
% اللغة والاتّجاه (polyglossia + bidi مع محرّك XeLaTeX)
\usepackage{polyglossia}
\setmainlanguage[numerals=maghrib]{arabic}
\setotherlanguage{english}
\setotherlanguage{french}
\IfFontExistsTF{Amiri}
  {\newfontfamily\arabicfont[Script=Arabic]{Amiri}}
  {\newfontfamily\arabicfont[Script=Arabic]{FreeSerif}}
\newfontfamily\englishfont{Latin Modern Roman}
\newcommand{\en}[1]{\textenglish{#1}}
\begin{document}
\sloppy
"""

COVER=r"""% ------------------------------ الغلاف (كما في الأصل) ---------------
\thispagestyle{empty}
\begin{titlepage}
\centering
{\large\bfseries شعبة علم النفس الإكلينيكي}\\[10pt]
{\large جامعة عبد المالك السعدي}\\[3pt]
{\large كلية الآداب والعلوم الإنسانية بتطوان}\\[36pt]
{\large بحث لنيل شهادة الإجازة في علم النفس الإكلينيكي}\\[16pt]
{\large\bfseries بعنوان:}\\[18pt]
{\LARGE\bfseries إدراك أسلوب المعاملة الوالدية المتسلط\\[10pt]
وعلاقته بتقدير الذات لدى المراهقين (15--17 سنة)}\\[10pt]
{\Large دراسة كيفية}\\[46pt]
{\large إعداد الطالبة: لطيفة أديب}\\[6pt]
{\large رقم التسجيل: 23068742}\\[16pt]
{\large تحت إشراف الأستاذ: د. بدر الدين الزيدي}\\[26pt]
{\large المستوى: السنة الثالثة من الإجازة -- علم النفس الإكلينيكي}\\[36pt]
{\large السنة الجامعية: 2025 -- 2026}
\end{titlepage}
\pagenumbering{roman}
"""

# شكر و ملخص و Résumé (from source, de-duplicated)
FRONT=r"""% ------------------------------ شكر وتقدير --------------------------
\chapter*{شكر وتقدير}
\addcontentsline{toc}{chapter}{شكر وتقدير}
\begin{center}\bfseries بسم الله الرحمن الرحيم\end{center}
الحمد لله تعالى الذي بنعمته تتمّ الصالحات، والذي وفّقني وأعانني على إنجاز هذا العمل، فله الحمد
أولًا وآخرًا.\par
أتقدّم بخالص الشكر وعظيم الامتنان إلى أستاذي المشرف الدكتور بدر الدين الزيدي، على ما أولاني من
توجيه علمي، ونصائح قيّمة، ومتابعة جادّة طوال مراحل إنجاز هذا البحث، فكان لتوجيهاته الأثر الكبير
في إخراج هذا العمل.\par
كما أتقدّم بجزيل الشكر إلى الدكتورة فوزية بلال، رئيسة شعبة علم النفس الإكلينيكي، على جهودها في
خدمة الطلبة، وعلى ما وفّرته من دعم أكاديمي أسهم في إنجاز هذا المسار الجامعي.\par
وأتوجّه بالشكر والتقدير إلى إدارة المؤسسة التعليمية التي احتضنت الجانب الميداني من الدراسة،
وإلى السيّد المدير، على تعاونه وتيسيره إجراءات إنجاز هذا البحث.\par
كما أتقدّم بخالص الشكر إلى جميع المراهقين المشاركين في هذه الدراسة، الذين أبدوا تعاونًا صادقًا،
وتقاسموا خبراتهم وتجاربهم بكلّ مسؤولية.\par
ولا يفوتني أن أعبّر عن عميق امتناني إلى زوجي، الذي كان خير سندٍ وداعمٍ لي طوال هذه الرحلة
العلمية، بصبره وتشجيعه ومساندته المستمرّة. كما أتوجّه بالشكر إلى أبنائي الأعزّاء، الذين تحمّلوا
انشغالي خلال فترة إعداد هذا البحث، فكان صبرهم وتفهّمهم دافعًا للاستمرار حتى إتمام هذا العمل.\par
وفي الختام، أتقدّم بالشكر والامتنان إلى كلّ من ساهم، من قريب أو بعيد، في إنجاز هذا البحث، سائلًا
الله تعالى أن يجزي الجميع خير الجزاء. والله وليّ التوفيق.

% ------------------------------ الملخّص -----------------------------
\chapter*{ملخّص الدراسة}
\addcontentsline{toc}{chapter}{ملخّص الدراسة}
هدفت هذه الدراسة إلى استكشاف إدراك أسلوب المعاملة الوالدية المتسلط وعلاقته بتقدير الذات لدى
المراهقين (15--17 سنة)، من خلال فهم الخبرة الذاتية للمراهقين كما يعبّرون عنها، والكشف عن الكيفية
التي يدركون بها الممارسات الوالدية المتسلطة، وكيف يفسّرون انعكاسها على مختلف أبعاد تقدير الذات
لديهم. وقد انطلقت الدراسة من ملاحظة أنّ معظم الدراسات السابقة تناولت العلاقة بين المفهومين من
خلال مقاربات كمّية، في حين ظلّ فهم المعاني التي يمنحها المراهق لخبراته الأسرية أقلّ تناولًا،
خاصّة في البيئة المغربية.\par
اعتمدت الدراسة المنهج الكيفي، واختيرت عيّنة قصدية متجانسة مكوّنة من أحد عشر مراهقًا ومراهقة،
تتراوح أعمارهم بين (15--17 سنة). ولجمع المعطيات استُخدمت المقابلة شبه الموجّهة، ثمّ حُلّلت
البيانات باستخدام تحليل المضمون الموجَّه.\par
أظهرت النتائج أنّ إدراك المراهقين لأسلوب المعاملة الوالدية المتسلط تمثّل في مجموعة من الممارسات،
أبرزها الضبط والصرامة، والعقاب، وغياب الحوار والتفسير، والبرود العاطفي وقلّة التقبّل، والرقابة
والتدخّل، وكبت الاستقلالية والرأي. كما بيّنت النتائج أنّ هذا الإدراك ارتبط بتباين أبعاد تقدير
الذات؛ إذ بدا تأثيره أكثر وضوحًا في البعدين الشخصي والأسري، وامتدّ لدى بعض المشاركين إلى البعدين
الاجتماعي والجسمي، في حين لم يتأثّر البعد الأكاديمي لدى بعض المراهقين، الذين فسّروا المتابعة
الدراسية الصارمة على أنّها تعبير عن حرص الوالدين على نجاحهم لا مجرّد مظهر من مظاهر التسلط.
وكشفت الدراسة أيضًا عن مؤشّرات نوعية، من أهمّها أثر الصراعات الأسرية، والتناقض بين أقوال الوالدين
وممارساتهما، واختلاف إدراك الإخوة للممارسات الوالدية داخل الأسرة الواحدة.\par
\noindent\textbf{الكلمات المفتاحية:} إدراك أسلوب المعاملة الوالدية المتسلط؛ تقدير الذات؛
المراهقون (15--17 سنة)؛ الدراسة الكيفية؛ تحليل المضمون الموجَّه.

% ------------------------------ Résumé ------------------------------
\chapter*{Résumé}
\addcontentsline{toc}{chapter}{Résumé}
\begin{french}
Cette étude qualitative vise à explorer la perception du style parental autoritaire et sa
relation avec l'estime de soi chez les adolescents âgés de 15 à 17 ans. Elle cherche à
comprendre la manière dont les adolescents perçoivent les pratiques parentales autoritaires
et comment ces expériences sont liées à leur estime de soi.\par
L'étude s'appuie sur une approche qualitative. Un échantillon raisonné et homogène composé de
onze adolescents a été sélectionné selon des critères précis. Les données ont été recueillies
au moyen d'entretiens semi-directifs, puis analysées à l'aide de l'analyse de contenu dirigée.\par
Les résultats montrent que les adolescents perçoivent le style parental autoritaire à travers
plusieurs pratiques récurrentes, notamment la discipline rigide, les sanctions, l'absence de
dialogue et d'explication, la froideur affective, le contrôle excessif ainsi que la limitation
de l'autonomie. Les effets apparaissent plus marqués au niveau des dimensions personnelle et
familiale, tandis que les dimensions sociale et physique varient selon les participants. En
revanche, la dimension académique n'a pas été affectée chez certains adolescents, qui ont
interprété le suivi scolaire rigoureux comme une marque d'intérêt des parents pour leur
réussite. L'étude met également en évidence plusieurs thèmes émergents : les conflits
familiaux, l'incohérence entre les paroles et les comportements des parents, ainsi que les
différences de perception entre frères et sœurs.\par
\noindent\textbf{Mots-clés :} Perception du style parental autoritaire ; Estime de soi ;
Adolescents (15--17 ans) ; Étude qualitative ; Analyse de contenu dirigée.
\end{french}

\cleardoublepage
\tableofcontents
\listoftables
\cleardoublepage
\pagenumbering{arabic}
"""

with open('/home/user/MP3Translator/rapport/_intro.tex',encoding='utf-8') as f:
    NEW_INTRO=f.read()
with open('/home/user/MP3Translator/rapport/_ch4.tex',encoding='utf-8') as f:
    NEW_CH4=f.read()

# ----------------------------------------------------------------------
# 6. Main conversion loop
# ----------------------------------------------------------------------
out=[PREAMBLE, COVER, FRONT]
i=0
n=len(blocks)
pending_caption=None
in_itemize=False
mode='skip_intro'   # skip old general intro until Chapter 1
emitted_ch2=False
emitted_ch4=False
in_appendix=False
refs_mode=False

def latin_counts(s):
    la=sum(1 for c in s if 'a'<=c.lower()<='z')
    ar=sum(1 for c in s if '؀'<=c<='ۿ')
    return la,ar

def close_itemize():
    global in_itemize
    if in_itemize:
        out.append(r'\end{itemize}')
        in_itemize=False

def is_table_caption(t):
    return t.strip().startswith('جدول (') or (t.strip().startswith('جدول ') and ':' in t)

def strip_caption(t):
    t=t.strip()
    if ':' in t:
        return t.split(':',1)[1].strip()
    return t

while i<n:
    kind,st,val=blocks[i]
    if kind=='T':
        close_itemize()
        out.append(emit_table(val, pending_caption))
        pending_caption=None
        i+=1; continue
    t=val.strip()

    # --- skip old general introduction (68-86) until Chapter 1 ---
    if mode=='skip_intro':
        if t.startswith('الفصل الاول') or t.startswith('الفصل الأول'):
            mode='normal'
            close_itemize()
            out.append(NEW_INTRO)
            out.append(r'\chapter{الإطار النظري للدراسة}')
            i+=1; continue
        i+=1; continue

    # --- old Chapter 4 -> replace with rebuilt version ---
    if t.startswith('الفصل الرابع'):
        close_itemize()
        out.append(NEW_CH4)
        emitted_ch4=True
        # skip everything until الخاتمة العامة
        i+=1
        while i<n:
            k2,s2,v2=blocks[i]
            if k2=='P' and v2.strip().startswith('الخاتمة العامة'):
                break
            i+=1
        continue

    # --- inject Chapter 2 heading before aims/problematic ---
    if (not emitted_ch2) and (t.startswith('4. هدف البحث') or t.startswith('4.هدف')
                              or ('هدف البحث' in t and len(t)<25)):
        close_itemize()
        out.append(r'\chapter{الإشكالية وإجراءات الدراسة المنهجية}')
        out.append(r'\section*{تمهيد}')
        out.append(r"""تتناول الدراسة الحالية العلاقة بين إدراك المراهق لأسلوب المعاملة الوالدية
المتسلط وتقديره لذاته، انطلاقًا من أنّ أثر الممارسات الوالدية لا يتحدّد بوجودها في ذاتها، وإنما
بالكيفية التي يدركها بها الأبناء ويفسّرونها. وقد اتّجهت الدراسة إلى المقاربة الكيفية باعتبارها
الأقدر على استكشاف هذه الخبرة، والكشف عن المعاني التي يبنيها المراهق انطلاقًا من علاقته بوالديه،
خاصّة في ظلّ ندرة الدراسات الكيفية في هذا الموضوع بالبيئة المغربية.""")
        out.append(r'\section{هدف البحث}')
        emitted_ch2=True
        i+=1; continue

    # --- fragment lines / garbled smartart : skip ---
    if t.startswith('ؤ ') or t.startswith('ؤمن'):
        i+=1; continue
    if 'الإطار النظري للدراسةالإطار' in val or 'تحديد المحور الأول وأبعاده (أسلوب' in val \
       or 'إجراء المقابلات وجمع البياناتإجراء' in val or 'تحليل المضمون (التكرار – الاتجاه العام – الحالات الخاصة)تحليل' in val:
        # clean replacement for the SmartArt pipeline (emit once)
        if 'PIPELINE_DONE' not in out:
            close_itemize()
            out.append(r'\begin{itemize}[leftmargin=1.6em,itemsep=1pt]')
            for step in ['بناء الإطار النظري للدراسة',
                         'تحديد المحور الأول وأبعاده (أسلوب المعاملة الوالدية المتسلط)',
                         'تحديد المحور الثاني وأبعاده (تقدير الذات)',
                         'بناء دليل المقابلة شبه الموجّهة',
                         'اختيار العيّنة القصدية وإجراء الفرز القبلي',
                         'إجراء المقابلات وجمع البيانات',
                         'تفريغ المقابلات والقراءة المتكرّرة لها',
                         'ترميز البيانات وتصنيف الترميزات داخل الأبعاد المحدّدة مسبقًا',
                         'تحليل المضمون (التكرار -- الاتجاه العام -- الحالات الخاصة)',
                         'عرض النتائج وتفسيرها']:
                out.append(r'\item %s'%step)
            out.append(r'\end{itemize}')
            out.append('% PIPELINE_DONE')
        i+=1; continue

    # --- table caption paragraph ---
    if is_table_caption(t):
        close_itemize()
        pending_caption=strip_caption(t)
        i+=1; continue
    # appendix caption "الملحق (x): جدول (x-1): TITLE"
    if t.startswith('الملحق (') and 'جدول' in t:
        close_itemize()
        pending_caption=strip_caption(t.split('جدول',1)[1] if 'جدول' in t else t)
        # also emit appendix label as a subheading
        lbl=t.split(':',1)[0].strip()
        out.append(r'\subsection*{%s}'%esc(lbl))
        i+=1; continue

    # --- chapters explicitly present ---
    if t.startswith('الفصل الثالث'):
        close_itemize()
        out.append(r'\chapter{عرض النتائج وتحليلها}')
        i+=1; continue
    if t.startswith('الخاتمة العامة'):
        close_itemize()
        out.append(r'\chapter*{الخاتمة العامة}')
        out.append(r'\addcontentsline{toc}{chapter}{الخاتمة العامة}')
        i+=1; continue
    if t.startswith('الملاحق') and len(t)<12:
        close_itemize()
        out.append(r'\appendix')
        out.append(r'\chapter{الملاحق}')
        in_appendix=True
        i+=1; continue
    if t.startswith('قائمة المراجع'):
        close_itemize()
        out.append(r'\chapter*{قائمة المراجع}')
        out.append(r'\addcontentsline{toc}{chapter}{قائمة المراجع}')
        refs_mode=True
        i+=1; continue
    if t.startswith('لائحة الجداول'):
        i+=1; continue   # handled by \listoftables

    # --- participant case header in appendix ---
    import re as _re
    if _re.match(r'^المشارك\s+م\d+', t) or _re.match(r'^ملخص الحالة', t):
        close_itemize()
        out.append(r'\subsection*{%s}'%esc(t))
        i+=1; continue
    # --- participant identity line (concatenated runs) ---
    if t.startswith('الرمز:'):
        close_itemize()
        e=esc(t)
        for k in ['الجنس:','العمر:','المستوى الدراسي:','تاريخ المقابلة:','عدد الإخوة:']:
            e=e.replace(k, r'\quad '+k)
        out.append(r'\noindent\textbf{%s}\par'%e)
        i+=1; continue

    # --- headings ---
    lvl,htext=heading_level(t)
    if lvl:
        close_itemize()
        if in_appendix:
            out.append(r'\subsection*{%s}'%esc(htext))
        else:
            cmd={'sec':r'\section','sub':r'\subsection','secstar':r'\section*'}[lvl]
            out.append(r'%s{%s}'%(cmd,esc(htext)))
        i+=1; continue

    # --- list paragraphs ---
    if st=='Paragraphedeliste':
        if not in_itemize:
            out.append(r'\begin{itemize}[leftmargin=1.6em,itemsep=1pt]')
            in_itemize=True
        if t.endswith(':') and len(t)<40:
            out.append(r'\item \textbf{%s}'%esc(t))
        else:
            out.append(r'\item %s'%esc(t))
        i+=1; continue
    else:
        close_itemize()

    # --- run-in bold short labels ---
    if (t.endswith(':') and len(t)<45) or any(t.startswith(k) for k in RUNIN):
        # split "label: rest"
        if ':' in t and not t.endswith(':'):
            lab,rest=t.split(':',1)
            out.append(r'\noindent\textbf{%s:} %s\par'%(esc(lab.strip()),esc(rest.strip())))
        else:
            out.append(r'\noindent\textbf{%s}\par'%esc(t))
        i+=1; continue

    # --- default paragraph ---
    if t.startswith('- '):
        t=t[2:].strip()
    la,ar=latin_counts(t)
    if la>ar and la>3:
        # Latin-dominant line (e.g. foreign reference) -> LTR
        body=r'\textenglish{%s}'%esc(t)
        if refs_mode:
            out.append(r'\noindent\hangindent=1.8em\hangafter=1 %s\par\vspace{3pt}'%body)
        else:
            out.append(r'\noindent %s\par'%body)
    elif refs_mode:
        out.append(r'\noindent\hangindent=1.8em\hangafter=1 %s\par\vspace{3pt}'%esc(t))
    else:
        out.append(esc(t))
        out.append('')   # blank line = paragraph break
    i+=1

close_itemize()
out.append(r'\end{document}')

with open(OUT,'w',encoding='utf-8') as f:
    f.write('\n'.join(out))
print("WROTE", OUT, "blocks:", n)
