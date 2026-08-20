from __future__ import annotations

from dataclasses import dataclass, field
from fractions import Fraction
import math
import random
import statistics
import re
from typing import Callable

import sympy as sp

try:
    from learning_outcomes_data import LEARNING_OUTCOMES
except Exception:
    LEARNING_OUTCOMES = {}

from syllabus_topics import offline_topic_rows, canonical_track_code


TRACKS = {
    "O-Level Mathematics (4052)": "O",
    "N(A)-Level Mathematics (4045)": "NA",
    "N(T)-Level Mathematics (4046)": "NT",
}


@dataclass(frozen=True)
class Topic:
    code: str
    name: str
    strand: str
    keywords: tuple[str, ...] = ()


@dataclass
class Question:
    track: str
    topic_code: str
    topic_name: str
    strand: str
    difficulty: str
    prompt: str
    target_skill: str
    hints: list[str]
    answer_display: str
    worked_solution: list[str]
    answer_kind: str = "numeric"
    answer_value: object = None
    family: str = ""
    learning_outcome_source: str = ""
    statistics_graph: dict | None = None


@dataclass
class StepFeedback:
    line_number: int
    line: str
    status: str
    feedback: str


@dataclass
class AttemptResult:
    is_correct: bool
    answer_score: int
    reasoning_score: int
    mastery: str
    summary: str
    first_logic_break: int | None = None
    first_logic_break_explanation: str = ""
    step_feedback: list[StepFeedback] = field(default_factory=list)
    strengths: list[str] = field(default_factory=list)
    gaps: list[str] = field(default_factory=list)
    next_hint: str = ""


TOPICS = {}
for _track in ("NT", "NA", "O", "G2A", "G3A"):
    _rows = offline_topic_rows(_track)
    TOPICS[_track] = [
        Topic(
            str(row["code"]),
            str(row["name"]),
            str(row["strand"]),
            tuple(
                dict.fromkeys(
                    [
                        str(row["name"]),
                        *re.findall(r"[A-Za-z][A-Za-z -]{3,}", str(row.get("details", "")))[:8],
                    ]
                )
            ),
        )
        for row in _rows
    ]


def topics_for_track(track: str) -> list[Topic]:
    return list(TOPICS.get(track, TOPICS.get("O", [])))


def official_topic_code(track: str, code: str) -> str:
    return code



OUTCOME_FOCI = {}
for _track in ("NT", "NA", "O", "G2A", "G3A"):
    OUTCOME_FOCI[_track] = {
        str(row["code"]): list(row.get("outcomes", []))
        for row in offline_topic_rows(_track)
    }


def _level_for_track(track: str) -> str:
    return {"NT": "G1", "NA": "G2", "O": "G3", "G2A": "G2"}.get(track, "G3")


def _learning_outcome(topic: Topic, track: str, rng: random.Random) -> tuple[str, str]:
    focuses = OUTCOME_FOCI.get(track, {}).get(topic.code, [])
    if focuses:
        return rng.choice(focuses), "learning outcomes(1).xlsx"
    return (f"Apply {topic.name.lower()} accurately in an appropriate problem.", "learning outcomes(1).xlsx")


def _topic(track: str, code: str) -> Topic:
    for t in topics_for_track(track):
        if t.code == code:
            return t
    return topics_for_track(track)[0]


def _q(track, topic, difficulty, prompt, skill, hints, answer, solution, kind="numeric", family="", source="", statistics_graph=None):
    return Question(
        track, topic.code, topic.name, topic.strand, difficulty,
        prompt, skill, hints, answer, solution,
        kind, answer, family, source, statistics_graph
    )


def _fmt_num(x):
    if isinstance(x, Fraction):
        return str(x.numerator) if x.denominator == 1 else f"{x.numerator}/{x.denominator}"
    if isinstance(x, float):
        if abs(x-round(x)) < 1e-10: return str(int(round(x)))
        return f"{x:.4g}"
    return str(x)


def _numbers(track, topic, diff, rng, skill, source):
    family_bank = {
        "Foundation": ["integer", "fractions", "rounding"] if track == "NT" else ["integer", "rounding", "factors"],
        "Similar": ["integer", "standard_form", "factors"] if track != "NT" else ["integer", "fractions", "rounding"],
        "Stretch": ["standard_form", "factors", "integer"] if track != "NT" else ["fractions", "integer", "rounding"],
    }
    fam = rng.choice(family_bank.get(diff, family_bank["Similar"]))
    if fam == "integer":
        if diff == "Foundation":
            a,b = rng.randint(-18,18), rng.randint(-12,12)
            ans=a-b
            return _q(track,topic,diff,f"Calculate {a} - ({b}).",skill,["Use the number line or sign rules."],str(ans),[f"{a}-({b})={ans}"],family=fam,source=source)
        if diff == "Similar":
            a,b,c=rng.randint(-12,18),rng.randint(-9,12),rng.randint(2,7)
            ans=a+b*c
            return _q(track,topic,diff,f"Evaluate {a} + {b} \\times {c}.",skill,["Apply multiplication before addition."],str(ans),[f"{b}\\times {c}={b*c}",f"{a}+{b*c}={ans}"],family=fam,source=source)
        a,b,c=rng.randint(12,30),rng.randint(-15,-3),rng.randint(3,8)
        ans=(a+b)*c
        return _q(track,topic,diff,f"A lift starts at floor {a}, moves {abs(b)} floors down, then repeats this final change {c} times in total. What floor number is represented by ({a}+({b}))\\times {c}?",skill,["Interpret the signed change first."],str(ans),[f"{a}+({b})={a+b}",f"({a+b})\\times {c}={ans}"],family=fam,source=source)
    if fam == "fractions":
        p,q=rng.randint(1,7),rng.randint(2,9); r,s=rng.randint(1,7),rng.randint(2,9)
        ans=Fraction(p,q)+Fraction(r,s)
        prompt=f"Calculate {p}/{q} + {r}/{s}." if diff!="Stretch" else f"A recipe uses {p}/{q} cup of milk and {r}/{s} cup of yoghurt. Find the total amount used."
        return _q(track,topic,diff,prompt,skill,["Use a common denominator."],_fmt_num(ans),[f"Use denominator {math.lcm(q,s)}.",f"Answer = {_fmt_num(ans)}"],kind="fraction",family=fam,source=source)
    if fam == "rounding":
        x=rng.uniform(100,9999)
        dp=1 if diff=="Foundation" else 2
        ans=round(x,dp)
        return _q(track,topic,diff,f"Round {x:.4f} to {dp} decimal place{'s' if dp>1 else ''}.",skill,["Check the digit immediately after the required place."],str(ans),[f"Rounded value = {ans}"],family=fam,source=source)
    if fam == "standard_form":
        a=rng.randint(12,98)/10; n=rng.randint(-5,5)
        if diff=="Foundation":
            val=a*(10**n)
            return _q(track,topic,diff,f"Write {a} \\times 10^{n} as an ordinary number.",skill,["Move the decimal point according to the power of 10."],_fmt_num(val),[f"{a}\\times10^{n}={_fmt_num(val)}"],family=fam,source=source)
        b=rng.randint(12,89)/10; m=rng.randint(-3,4)
        val=a*b*10**(n+m)
        return _q(track,topic,diff,f"Calculate ({a} \\times 10^{n})({b} \\times 10^{m}). Give your answer in standard form.",skill,["Multiply the coefficients and add the powers."],f"{val:.6g}",[f"{a}\\times {b}={a*b:.4g}",f"10^{n}\\times10^{m}=10^{n+m}"],family=fam,source=source)
    # factors
    a=rng.choice([24,36,48,60,72,84,90,96]); b=rng.choice([30,42,54,66,78,90])
    g=math.gcd(a,b); l=abs(a*b)//g
    if diff=="Foundation":
        return _q(track,topic,diff,f"Find the HCF of {a} and {b}.",skill,["Write each number as a product of prime factors."],str(g),[f"HCF({a},{b})={g}"],family=fam,source=source)
    return _q(track,topic,diff,f"Two lights flash every {a} seconds and {b} seconds. They flash together now. After how many seconds will they next flash together?",skill,["This is an LCM problem."],str(l),[f"LCM({a},{b})={l}"],family=fam,source=source)


def _ratio(track,topic,diff,rng,skill,source):
    family_bank={
        "Foundation":["equivalent","share"],
        "Similar":["share","equivalent","map"],
        "Stretch":["share","map"],
    }
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    if fam=="share":
        a,b=rng.randint(2,7),rng.randint(2,8); b = b+1 if b==a else b; unit=rng.randint(8,25); total=(a+b)*unit
        if diff=="Foundation":
            ans=a*unit
            prompt=f"Divide ${total} in the ratio {a}:{b}. Find the smaller share." if a<=b else f"Divide ${total} in the ratio {a}:{b}. Find the first share."
        elif diff=="Similar":
            ans=b*unit
            prompt=f"A prize of ${total} is shared between A and B in the ratio {a}:{b}. Find B's share."
        else:
            ans=total
            first=a*unit
            prompt=f"A sum of money is shared in the ratio {a}:{b}. The first person receives ${first}. Find the total sum."
        return _q(track,topic,diff,prompt,skill,["Find the value of one ratio part first."],str(ans),[f"One part = {unit}",f"Required amount = {ans}"],family=fam,source=source)
    if fam=="equivalent":
        a,b=rng.randint(2,9),rng.randint(2,9); b = b+1 if b==a and b<9 else (b-1 if b==a else b); k=rng.randint(2,9)
        if diff=="Foundation":
            return _q(track,topic,diff,f"Complete the equivalent ratio {a}:{b} = {a*k}:x.",skill,["Both terms are multiplied by the same factor."],str(b*k),[f"Scale factor = {k}",f"x={b*k}"],family=fam,source=source)
        x=a*k
        return _q(track,topic,diff,f"The ratio of red to blue beads is {a}:{b}. There are {x} red beads. How many blue beads are there?",skill,["Use equivalent ratios."],str(b*k),[f"Scale factor={k}",f"Blue={b*k}"],family=fam,source=source)
    scale=rng.choice([20000,25000,50000,100000]); mapcm=rng.randint(2,9); real_cm=scale*mapcm; km=real_cm/100000
    prompt=f"A map has scale 1:{scale}. Two places are {mapcm} cm apart on the map. Find the actual distance in km."
    return _q(track,topic,diff,prompt,skill,["Convert the scale distance to actual centimetres, then to kilometres."],_fmt_num(km),[f"Actual distance={real_cm} cm",f"={_fmt_num(km)} km"],family=fam,source=source)


def _percentage(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["discount"],"Similar":["discount","change"],"Stretch":["reverse","change"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    pct=rng.choice([5,10,12,15,20,25,30]); base=rng.choice([80,120,160,200,240,320,500])
    if fam=="discount":
        ans=base*(100-pct)/100
        return _q(track,topic,diff,f"An item costs ${base}. It is discounted by {pct}%. Find the sale price.",skill,["Find the discount or multiply by the remaining percentage."],_fmt_num(ans),[f"Sale price={base}\\times {(100-pct)/100}",f"=${_fmt_num(ans)}"],family=fam,source=source)
    if fam=="reverse":
        final=base*(100+pct)/100
        return _q(track,topic,diff,f"After an increase of {pct}%, a quantity is {_fmt_num(final)}. Find its original value.",skill,["The final value represents more than 100%."],str(base),[f"{100+pct}% corresponds to {_fmt_num(final)}",f"100%={base}"],family=fam,source=source)
    new=base+rng.choice([20,40,60,80]); ans=(new-base)/base*100
    return _q(track,topic,diff,f"A value increases from {base} to {new}. Find the percentage increase.",skill,["Percentage change = change ÷ original × 100%."],_fmt_num(ans),[f"Change={new-base}",f"Percentage increase={_fmt_num(ans)}%"],family=fam,source=source)


def _rate(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["unit_rate","reverse"],"Similar":["speed","reverse"],"Stretch":["speed"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    if fam=="speed":
        t=rng.choice([1.5,2,2.5,3]); speed=rng.choice([40,50,60,72,80]); d=t*speed
        if diff=="Stretch":
            d2=rng.choice([30,45,60]); t2=rng.choice([0.5,0.75,1]); ans=(d+d2)/(t+t2)
            prompt=f"A vehicle travels {_fmt_num(d)} km in {t} h, then {d2} km in {t2} h. Find its average speed in km/h."
        else:
            ans=speed; prompt=f"A vehicle travels {_fmt_num(d)} km in {t} hours. Find its average speed in km/h."
        return _q(track,topic,diff,prompt,skill,["Average speed = total distance ÷ total time."],_fmt_num(ans),[f"Average speed = {_fmt_num(ans)} km/h"],family=fam,source=source)
    if fam=="unit_rate":
        qty=rng.randint(3,9); price=qty*rng.choice([2.5,3,4,5,6]); ans=price/qty
        return _q(track,topic,diff,f"{qty} identical notebooks cost ${_fmt_num(price)}. Find the cost of one notebook.",skill,["Divide total cost by quantity."],_fmt_num(ans),[f"Unit cost={_fmt_num(price)}/{qty}={_fmt_num(ans)}"],family=fam,source=source)
    speed=rng.choice([45,60,75,90]); time=rng.choice([1.2,1.5,2.4]); ans=speed*time
    return _q(track,topic,diff,f"A car travels at {speed} km/h for {time} h. Find the distance travelled.",skill,["Distance = speed × time."],_fmt_num(ans),[f"Distance={speed}\\times{time}={_fmt_num(ans)} km"],family=fam,source=source)


def _algebra(track,topic,diff,rng,skill,source):
    family_bank={
        "Foundation":["simplify","substitute","expand"],
        "Similar":["expand","substitute","formula","simplify"],
        "Stretch":["formula","expand","substitute"],
    }
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    x=sp.symbols('x')
    if fam=="simplify":
        a,b,c=rng.randint(2,7),rng.randint(1,6),rng.randint(1,6)
        expr=a*x+b*x-c*x; ans=sp.expand(expr)
        return _q(track,topic,diff,f"Simplify {a}x + {b}x - {c}x.",skill,["Collect like terms."],sp.sstr(ans),[f"({a}+{b}-{c})x={sp.sstr(ans)}"],kind="expr",family=fam,source=source)
    if fam=="expand":
        a,b,c=rng.randint(2,5),rng.randint(-6,6),rng.randint(-6,6)
        if diff=="Foundation": expr=a*(x+b); ans=sp.expand(expr); prompt=f"Expand {a}(x {'+' if b>=0 else '-'} {abs(b)})."
        else: expr=(x+b)*(x+c); ans=sp.expand(expr); prompt=f"Expand and simplify (x {'+' if b>=0 else '-'} {abs(b)})(x {'+' if c>=0 else '-'} {abs(c)})."
        return _q(track,topic,diff,prompt,skill,["Multiply every term in one bracket by every term in the other."],sp.sstr(ans),[f"Result: {sp.sstr(ans)}"],kind="expr",family=fam,source=source)
    if fam=="substitute":
        a,b,c=rng.randint(2,6),rng.randint(-5,5),rng.randint(-8,8); xv=rng.randint(-4,6); ans=a*xv*xv+b*xv+c
        return _q(track,topic,diff,f"Evaluate {a}x^2 {'+' if b>=0 else '-'} {abs(b)}x {'+' if c>=0 else '-'} {abs(c)} when x = {xv}.",skill,["Substitute the value using brackets."],str(ans),[f"Substitute x={xv}",f"Answer={ans}"],family=fam,source=source)
    # formula
    a,b=rng.randint(2,8),rng.randint(2,10); xval=rng.randint(2,9); y=a*xval+b
    if diff=="Stretch":
        return _q(track,topic,diff,f"The formula y = {a}x + {b}. Given y = {y}, find x.",skill,["Rearrange the formula to make x the subject."],str(xval),[f"{a}x={y-b}",f"x={xval}"],family=fam,source=source)
    return _q(track,topic,diff,f"Use y = {a}x + {b} to find y when x = {xval}.",skill,["Substitute x into the formula."],str(y),[f"y={a}({xval})+{b}={y}"],family=fam,source=source)


def _equations(track,topic,diff,rng,skill,source):
    family_bank={
        "Foundation":["linear"],
        "Similar":["linear","word"],
        "Stretch":["word","inequality"] if track != "NT" else ["word","linear"],
    }
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    x=sp.symbols('x')
    a=rng.randint(2,8); sol=rng.randint(2,12); b=rng.randint(-8,10); c=a*sol+b
    if fam=="linear":
        if diff=="Foundation": prompt=f"Solve {a}x = {a*sol}."; steps=[f"x={sol}"]
        else: prompt=f"Solve {a}x {'+' if b>=0 else '-'} {abs(b)} = {c}."; steps=[f"{a}x={c-b}",f"x={sol}"]
        return _q(track,topic,diff,prompt,skill,["Undo the operations while keeping both sides balanced."],str(sol),steps,family=fam,source=source)
    if fam=="word":
        price=a; fixed=max(1,b+10); total=price*sol+fixed
        prompt=f"A taxi fare is a fixed ${fixed} plus ${price} per kilometre. The total fare is ${total}. Find the distance travelled."
        return _q(track,topic,diff,prompt,skill,["Form a linear equation for the total fare."],str(sol),[f"{price}x+{fixed}={total}",f"x={sol}"],family=fam,source=source)
    bound=rng.randint(3,12); aa=rng.randint(2,5); rhs=aa*bound+rng.randint(1,aa-1)
    ans=f"x < {math.ceil(rhs/aa)}"
    return _q(track,topic,diff,f"Solve the inequality {aa}x < {rhs}, where x is an integer.",skill,["Divide both sides by the positive coefficient."],ans,[f"x < {rhs/aa:.3g}",ans],kind="text",family=fam,source=source)


def _functions(track,topic,diff,rng,skill,source):
    family_bank={
        "Foundation":["table"],
        "Similar":["table","linear"],
        "Stretch":["linear","quadratic"] if track=="O" else ["linear"],
    }
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    if fam=="table":
        a,b=rng.randint(2,5),rng.randint(-4,4); xv=rng.randint(-3,5); ans=a*xv+b
        return _q(track,topic,diff,f"For y = {a}x {'+' if b>=0 else '-'} {abs(b)}, find y when x = {xv}.",skill,["Treat x as the input and y as the output."],str(ans),[f"y={a}({xv})+({b})={ans}"],family=fam,source=source)
    if fam=="linear":
        m=rng.randint(-4,5) or 2; c=rng.randint(-6,6); x1=rng.randint(-3,2); x2=x1+rng.randint(2,5); y1=m*x1+c; y2=m*x2+c
        if diff=="Foundation":
            return _q(track,topic,diff,f"A straight line passes through ({x1},{y1}) and ({x2},{y2}). Find its gradient.",skill,["Gradient = change in y ÷ change in x."],str(m),[f"m=({y2}-{y1})/({x2}-{x1})={m}"],family=fam,source=source)
        return _q(track,topic,diff,f"Find the equation of the line with gradient {m} and y-intercept {c}.",skill,["Use y = mx + c."],f"y={m}x+{c}",[f"y={m}x+{c}"],kind="text",family=fam,source=source)
    h,k=rng.randint(-3,3),rng.randint(-5,5); ans=f"({-h},{k})"
    return _q(track,topic,diff,f"The quadratic function is y = (x {'+' if h>=0 else '-'} {abs(h)})^2 {'+' if k>=0 else '-'} {abs(k)}. State the coordinates of its turning point.",skill,["Use completed-square form."],ans,[f"Turning point = ({-h},{k})"],kind="text",family=fam,source=source)


def _sequence(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["nth"],"Similar":["nth","missing"],"Stretch":["reverse","missing"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    a=rng.randint(-5,12); d=rng.randint(2,8); seq=[a+d*i for i in range(5)]
    nth=f"{d}n {'+' if a-d>=0 else '-'} {abs(a-d)}"
    if fam=="nth":
        return _q(track,topic,diff,f"Find the nth term of the sequence {', '.join(map(str,seq))}, ...",skill,["Find the common difference, then compare with dn."],nth,[f"Common difference={d}",f"nth term={nth}"],kind="text",family=fam,source=source)
    if fam=="missing":
        k=rng.randint(6,15); ans=a+d*(k-1)
        return _q(track,topic,diff,f"The nth term of a sequence is {nth}. Find the {k}th term.",skill,["Substitute n into the nth-term formula."],str(ans),[f"Term={d}({k})+({a-d})={ans}"],family=fam,source=source)
    target=a+d*rng.randint(8,15); n=(target-(a-d))/d
    return _q(track,topic,diff,f"The nth term of a sequence is {nth}. Which term is equal to {target}?",skill,["Set the nth-term expression equal to the target value."],str(int(n)),[f"{d}n+({a-d})={target}",f"n={int(n)}"],family=fam,source=source)


def _geometry(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["angles","parallel"],"Similar":["angles","polygon","parallel"],"Stretch":["polygon","parallel"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    if fam=="angles":
        a=rng.randint(25,75); ans=180-a
        return _q(track,topic,diff,f"Two angles on a straight line are {a}° and x°. Find x.",skill,["Angles on a straight line sum to 180°."],str(ans),[f"x=180-{a}={ans}°"],family=fam,source=source)
    if fam=="polygon":
        n=rng.randint(5,10); ans=(n-2)*180
        return _q(track,topic,diff,f"Find the sum of the interior angles of a {n}-sided polygon.",skill,["Use (n−2)×180°."],str(ans),[f"({n}-2)\\times180={ans}°"],family=fam,source=source)
    a=rng.randint(35,75); ans=a
    return _q(track,topic,diff,f"Two parallel lines are cut by a transversal. One corresponding angle is {a}°. Find the corresponding angle on the other line.",skill,["Corresponding angles between parallel lines are equal."],str(ans),[f"Angle={a}°"],family=fam,source=source)


def _mensuration(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["rectangle"],"Similar":["rectangle","circle","volume"],"Stretch":["circle","volume"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    if fam=="rectangle":
        l,w=rng.randint(5,18),rng.randint(3,12)
        if diff=="Foundation": ans=2*(l+w); prompt=f"A rectangle is {l} cm by {w} cm. Find its perimeter."
        else: ans=l*w; prompt=f"A rectangular floor is {l} m by {w} m. Find its area."
        return _q(track,topic,diff,prompt,skill,["Choose the correct perimeter or area formula."],str(ans),[f"Answer={ans}"],family=fam,source=source)
    if fam=="circle":
        r=rng.randint(3,10); ans=math.pi*r*r
        return _q(track,topic,diff,f"A circle has radius {r} cm. Find its area in terms of π.",skill,["Area = πr²."],f"{r*r}pi",[f"Area=π({r})^2={r*r}π cm^2"],kind="text",family=fam,source=source)
    l,w,h=rng.randint(3,10),rng.randint(3,10),rng.randint(2,8); ans=l*w*h
    return _q(track,topic,diff,f"A cuboid measures {l} cm by {w} cm by {h} cm. Find its volume.",skill,["Volume = length × width × height."],str(ans),[f"V={l}\\times{w}\\times{h}={ans} cm^3"],family=fam,source=source)


def _pyth_trig(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["pyth","similar"],"Similar":["pyth","trig","similar"],"Stretch":["trig","similar"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    if fam=="pyth":
        triples=rng.choice([(3,4,5),(5,12,13),(8,15,17),(7,24,25)]); a,b,c=triples
        return _q(track,topic,diff,f"A right-angled triangle has perpendicular sides {a} cm and {b} cm. Find the hypotenuse.",skill,["Apply Pythagoras' theorem."],str(c),[f"c^2={a}^2+{b}^2={c*c}",f"c={c}"],family=fam,source=source)
    if fam=="trig":
        angle=rng.choice([30,35,40,45,50,60]); adj=rng.randint(5,15); opp=adj*math.tan(math.radians(angle))
        return _q(track,topic,diff,f"In a right-angled triangle, an angle is {angle}° and its adjacent side is {adj} cm. Find the opposite side, correct to 1 decimal place.",skill,["Use tan θ = opposite/adjacent."],f"{opp:.1f}",[f"opposite={adj}tan({angle}°)={opp:.1f} cm"],family=fam,source=source)
    k=rng.randint(2,5); small=rng.randint(3,8); big=small*k
    return _q(track,topic,diff,f"Two similar figures have scale factor {k} from small to large. A corresponding side on the small figure is {small} cm. Find the side on the large figure.",skill,["Corresponding lengths are in the scale-factor ratio."],str(big),[f"{small}\\times{k}={big}"],family=fam,source=source)


def _safe_mode(values):
    """Return a deterministic mode, or the smallest modal value if tied."""
    modes = statistics.multimode(values)
    if not modes:
        raise ValueError("Cannot determine a mode from empty data.")
    try:
        return min(modes)
    except TypeError:
        return modes[0]



def _cf_interpolate_x(xs, ys, target_y):
    target=float(target_y)
    for x1,y1,x2,y2 in zip(xs,ys,xs[1:],ys[1:]):
        y1=float(y1); y2=float(y2)
        if y1 <= target <= y2 and y2 > y1:
            return float(x1)+(target-y1)*(float(x2)-float(x1))/(y2-y1)
    return float(xs[-1])


def _cumulative_frequency_question(track, topic, diff, rng, skill, source):
    xs=[0,20,30,40,50,60,70,80,90,100]
    ys=[0,5,12,27,50,73,88,95,98,100]
    freqs=[ys[i+1]-ys[i] for i in range(len(ys)-1)]
    graph={
        "graph_type":"cumulative_frequency",
        "title":"Cumulative frequency of test scores",
        "x_label":"Score",
        "y_label":"Cumulative frequency",
        "class_boundaries":xs,
        "frequencies":freqs,
        "cumulative_frequencies":ys,
        "show_completed_graph_in_question":True,
        "show_grid":True,
    }

    if diff=="Foundation":
        family=rng.choice(["median","below"])
    elif diff=="Stretch":
        family=rng.choice(["iqr","percentile","above"])
    else:
        family=rng.choice(["median","quartile","below","above"])

    if family=="median":
        ans=round(_cf_interpolate_x(xs,ys,50))
        return _q(
            track,topic,diff,
            "The cumulative frequency curve shows the scores of 100 students. Use the graph to estimate the median score.",
            skill,
            ["For 100 students, locate cumulative frequency 50.","Read horizontally to the curve, then vertically to the score axis."],
            str(ans),
            ["N=100, so the median is at cumulative frequency 50.",f"Estimated median score = {ans}."],
            family=family,source=source,statistics_graph=graph
        )

    if family=="quartile":
        which=rng.choice(["lower","upper"])
        target=25 if which=="lower" else 75
        ans=round(_cf_interpolate_x(xs,ys,target))
        qname="lower quartile" if which=="lower" else "upper quartile"
        return _q(
            track,topic,diff,
            f"The cumulative frequency curve shows the scores of 100 students. Use the graph to estimate the {qname}.",
            skill,
            [f"Read the {qname} at cumulative frequency {target}.","Move horizontally to the curve, then vertically to the score axis."],
            str(ans),
            [f"For N=100, read the {qname} at cumulative frequency {target}.",f"Estimated {qname} score = {ans}."],
            family=family,source=source,statistics_graph=graph
        )

    if family=="iqr":
        q1=_cf_interpolate_x(xs,ys,25)
        q3=_cf_interpolate_x(xs,ys,75)
        ans=round(q3-q1)
        return _q(
            track,topic,diff,
            "The cumulative frequency curve shows the scores of 100 students. Use the graph to estimate the interquartile range.",
            skill,
            ["Read Q1 at cumulative frequency 25 and Q3 at cumulative frequency 75.","Calculate Q3 - Q1."],
            str(ans),
            [f"Q1 ≈ {round(q1)}, Q3 ≈ {round(q3)}.",f"IQR ≈ {round(q3)} - {round(q1)} = {ans}."],
            family=family,source=source,statistics_graph=graph
        )

    if family=="percentile":
        percentile=rng.choice([80,90])
        ans=round(_cf_interpolate_x(xs,ys,percentile))
        return _q(
            track,topic,diff,
            f"The cumulative frequency curve shows the scores of 100 students. Estimate the {percentile}th percentile score.",
            skill,
            [f"With 100 students, use cumulative frequency {percentile}.","Read across to the curve and then down to the score axis."],
            str(ans),
            [f"Read the curve at cumulative frequency {percentile}.",f"Estimated percentile score = {ans}."],
            family=family,source=source,statistics_graph=graph
        )

    threshold=rng.choice([40,60,70,80])
    cf=dict(zip(xs,ys))[threshold]
    if family=="above":
        ans=100-cf
        prompt=f"The cumulative frequency curve shows the scores of 100 students. Estimate the number of students who scored more than {threshold}."
        hints=[f"Read the cumulative frequency at score {threshold}.","Subtract it from 100."]
        solution=[f"At score {threshold}, cumulative frequency ≈ {cf}.",f"Number above {threshold} = 100 - {cf} = {ans}."]
    else:
        ans=cf
        prompt=f"The cumulative frequency curve shows the scores of 100 students. Estimate the number of students who scored {threshold} or less."
        hints=[f"Read the cumulative frequency directly above score {threshold}."]
        solution=[f"At score {threshold}, cumulative frequency ≈ {cf}.",f"So about {ans} students scored {threshold} or less."]

    return _q(track,topic,diff,prompt,skill,hints,str(ans),solution,family=family,source=source,statistics_graph=graph)


def _statistics(track,topic,diff,rng,skill,source):
    if "cumulative frequency" in topic.name.lower():
        return _cumulative_frequency_question(track,topic,diff,rng,skill,source)

    family_bank={"Foundation":["mean","median","range"],"Similar":["mean","median","range"],"Stretch":["mean","median","range"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    data=[rng.randint(3,20) for _ in range(rng.choice([5,6,7]))]
    if fam=="mean": ans=sum(data)/len(data); name="mean"
    elif fam=="median": ans=float(statistics.median(data)); name="median"
    else: ans=max(data)-min(data); name="range"
    prompt=f"Find the {name} of the data set: {', '.join(map(str,data))}."
    return _q(track,topic,diff,prompt,skill,[f"Use the definition of the {name}."],_fmt_num(ans),[f"{name.title()}={_fmt_num(ans)}"],family=fam,source=source)


def _probability(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["single"],"Similar":["single","complement"],"Stretch":["two_stage","complement"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    red,blue=rng.randint(2,8),rng.randint(2,8); total=red+blue
    if fam=="single": ans=Fraction(red,total); prompt=f"A bag contains {red} red and {blue} blue counters. One counter is chosen at random. Find P(red)."
    elif fam=="complement": ans=Fraction(blue,total); prompt=f"A bag contains {red} red and {blue} blue counters. One counter is chosen. Find the probability that it is not red."
    else:
        ans=Fraction(red,total)*Fraction(red-1,total-1) if red>1 else Fraction(0,1)
        prompt=f"A bag contains {red} red and {blue} blue counters. Two counters are chosen without replacement. Find the probability that both are red."
    return _q(track,topic,diff,prompt,skill,["Write favourable outcomes over total outcomes; for two stages multiply conditional probabilities."],_fmt_num(ans),[f"Probability={_fmt_num(ans)}"],kind="fraction",family=fam,source=source)


def _circle(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["radius"],"Similar":["angle","tangent"],"Stretch":["angle","tangent"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    if fam=="radius":
        d=rng.choice([8,10,12,14,18]); return _q(track,topic,diff,f"A circle has diameter {d} cm. Find its radius.",skill,["Radius is half the diameter."],str(d/2 if d%2 else d//2),[f"r={d}/2={d/2:g} cm"],family=fam,source=source)
    if fam=="angle":
        a=rng.randint(25,70); ans=2*a
        return _q(track,topic,diff,f"An angle at the circumference subtending an arc is {a}°. Find the angle at the centre subtending the same arc.",skill,["The angle at the centre is twice the angle at the circumference."],str(ans),[f"Central angle=2({a})={ans}°"],family=fam,source=source)
    return _q(track,topic,diff,"A radius is drawn to the point where a tangent touches a circle. Find the angle between the radius and the tangent.",skill,["A tangent is perpendicular to the radius at the point of contact."],"90",["Angle = 90°"],family=fam,source=source)



def _matrix_mathio(rows):
    """Return a MathIO/LaTeX matrix source using pmatrix."""
    body = r" \\ ".join(" & ".join(str(v) for v in row) for row in rows)
    return r"\begin{pmatrix}" + body + r"\end{pmatrix}"


def _column_vector_mathio(values):
    """Return a MathIO/LaTeX column vector."""
    return _matrix_mathio([[v] for v in values])


def _matrices(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["add","multiply_scalar"],"Similar":["add","product"],"Stretch":["product"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    a,b,c,d=[rng.randint(-5,8) for _ in range(4)]
    A=[[a,b],[c,d]]

    if fam=="add":
        e,f,g,h=[rng.randint(-5,8) for _ in range(4)]
        B=[[e,f],[g,h]]
        ans=[[a+e,b+f],[c+g,d+h]]
        prompt=(
            "Given "
            rf"\(A={_matrix_mathio(A)}\) and "
            rf"\(B={_matrix_mathio(B)}\), "
            rf"find \(A+B\)."
        )

    elif fam=="multiply_scalar":
        k=rng.randint(2,5)
        ans=[[k*a,k*b],[k*c,k*d]]
        prompt=(
            "Given "
            rf"\(A={_matrix_mathio(A)}\), "
            rf"find \({k}A\)."
        )

    else:
        x,y=rng.randint(1,6),rng.randint(1,6)
        v=[x,y]
        ans=[a*x+b*y,c*x+d*y]
        prompt=(
            "Calculate "
            rf"\({_matrix_mathio(A)}{_column_vector_mathio(v)}\)."
        )

    answer_math = (
        _matrix_mathio(ans)
        if isinstance(ans[0], list)
        else _column_vector_mathio(ans)
    )
    disp=rf"\({answer_math}\)"

    return _q(
        track,topic,diff,prompt,skill,
        ["Use the appropriate matrix operation row by row."],
        disp,
        [rf"Answer \(={answer_math}\)."],
        kind="text",family=fam,source=source
    )


def _vectors(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["component","magnitude"],"Similar":["component","position"],"Stretch":["position","magnitude"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    a,b=rng.randint(-6,8),rng.randint(-6,8)

    if fam=="component":
        c,d=rng.randint(-6,8),rng.randint(-6,8)
        u=[a,b]
        v=[c,d]
        ans=[a+c,b+d]
        prompt=(
            "Given "
            rf"\(\mathbf{{u}}={_column_vector_mathio(u)}\) and "
            rf"\(\mathbf{{v}}={_column_vector_mathio(v)}\), "
            rf"find \(\mathbf{{u}}+\mathbf{{v}}\)."
        )
        answer_math=_column_vector_mathio(ans)

    elif fam=="magnitude":
        a,b=rng.choice([(3,4),(5,12),(8,15)])
        u=[a,b]
        ans=int(math.hypot(a,b))
        prompt=(
            "Find the magnitude of "
            rf"\(\mathbf{{u}}={_column_vector_mathio(u)}\)."
        )
        answer_math=str(ans)

    else:
        c,d=rng.randint(-5,8),rng.randint(-5,8)
        ans=[c-a,d-b]
        prompt=(
            rf"Point \(A=({a},{b})\) and point \(B=({c},{d})\). "
            rf"Find \(\overrightarrow{{AB}}\) in column-vector form."
        )
        answer_math=_column_vector_mathio(ans)

    disp=rf"\({answer_math}\)"
    return _q(
        track,topic,diff,prompt,skill,
        ["Use vector components consistently."],
        disp,
        [rf"Answer \(={answer_math}\)."],
        kind="text",family=fam,source=source
    )



def _addmath_algebra(track, topic, diff, rng, skill, source):
    families={"Foundation":["expand","quadratic"],"Similar":["quadratic","surds","log"],"Stretch":["identity","reverse_quadratic","partial_fraction"]}
    fam=rng.choice(families.get(diff,families["Similar"]))
    x=sp.Symbol("x")
    if fam=="expand":
        a,b=rng.randint(2,5),rng.randint(1,7); ans=sp.expand((a*x+b)*(x-b))
        return _q(track,topic,diff,f"Expand and simplify ({a}x + {b})(x - {b}).",skill,["Expand each term and collect like terms."],str(ans),[str(ans)],kind="expr",family=fam,source=source)
    if fam in {"quadratic","reverse_quadratic"}:
        r1,r2=rng.sample(range(-6,7),2)
        if fam=="quadratic":
            expr=sp.expand((x-r1)*(x-r2)); ans=f"x = {r1} or x = {r2}"
            return _q(track,topic,diff,f"Solve {sp.sstr(expr)} = 0.",skill,["Factorise or use a suitable quadratic method."],ans,[f"(x-{r1})(x-{r2})=0",ans],kind="text",family=fam,source=source)
        s,p=r1+r2,r1*r2; ans=f"x^2 - ({s})x + ({p}) = 0"
        return _q(track,topic,diff,f"A quadratic has roots {r1} and {r2}. Form the monic quadratic equation.",skill,["Use sum and product of roots."],ans,[ans],kind="text",family=fam,source=source)
    if fam=="surds":
        n=rng.choice([8,12,18,20,27,32,45,50]); ans=sp.sqrt(n).simplify()
        return _q(track,topic,diff,f"Simplify sqrt({n}) fully.",skill,["Extract the largest square factor."],str(ans),[str(ans)],kind="text",family=fam,source=source)
    if fam=="log":
        base=rng.choice([2,3,5]); power=rng.randint(2,5); val=base**power
        return _q(track,topic,diff,f"Given log base {base} of {val} equals k, find k.",skill,["Rewrite in exponential form."],str(power),[f"{base}^{power}={val}",f"k={power}"],family=fam,source=source)
    if fam=="partial_fraction":
        a,b=rng.randint(1,4),rng.randint(1,4); expr=sp.apart((a*x+b)/((x+1)*(x+2)),x)
        return _q(track,topic,diff,f"Express ({a}x+{b})/((x+1)(x+2)) in partial fractions.",skill,["Use A/(x+1)+B/(x+2)."],str(expr),[str(expr)],kind="text",family=fam,source=source)
    a=rng.randint(2,5); ans=sp.expand((x+a)**2-x**2)
    return _q(track,topic,diff,f"Simplify (x+{a})^2 - x^2.",skill,["Expand and collect like terms."],str(ans),[str(ans)],kind="expr",family=fam,source=source)


def _addmath_trig(track, topic, diff, rng, skill, source):
    families={"Foundation":["exact","amplitude"],"Similar":["period","solve"],"Stretch":["identity","phase"]}
    fam=rng.choice(families.get(diff,families["Similar"]))
    if fam=="exact":
        angle=rng.choice([30,45,60]); fn=rng.choice(["sin","cos"])
        table={"sin":{30:"1/2",45:"sqrt(2)/2",60:"sqrt(3)/2"},"cos":{30:"sqrt(3)/2",45:"sqrt(2)/2",60:"1/2"}}
        ans=table[fn][angle]
        return _q(track,topic,diff,f"Find the exact value of {fn}({angle} degrees).",skill,["Use exact trigonometric values."],ans,[ans],kind="text",family=fam,source=source)
    if fam=="amplitude":
        a=rng.randint(2,6); d=rng.randint(-3,3)
        return _q(track,topic,diff,f"For y = {a} sin(x) + {d}, state the amplitude.",skill,["Amplitude is the absolute sine coefficient."],str(a),[f"Amplitude={a}"],family=fam,source=source)
    if fam=="period":
        b=rng.randint(2,5); ans=f"2pi/{b}"
        return _q(track,topic,diff,f"State the period of y = sin({b}x).",skill,["Period is 2pi/b."],ans,[ans],kind="text",family=fam,source=source)
    if fam=="solve":
        angle=rng.choice([30,45,60]); val={30:"1/2",45:"sqrt(2)/2",60:"sqrt(3)/2"}[angle]
        ans=f"x = {angle} or x = {180-angle}"
        return _q(track,topic,diff,f"Solve sin(x) = {val} for 0 <= x <= 180 degrees.",skill,["Use exact-angle knowledge or the sine graph."],ans,[ans],kind="text",family=fam,source=source)
    if fam=="phase":
        c=rng.choice([30,45,60]); ans=f"{c} degrees to the right"
        return _q(track,topic,diff,f"State the horizontal phase shift of y = sin(x - {c} degrees).",skill,["Compare with sin(x-c)."],ans,[ans],kind="text",family=fam,source=source)
    return _q(track,topic,diff,"Simplify sin(x)^2 + cos(x)^2.",skill,["Use the Pythagorean identity."],"1",["1"],family=fam,source=source)


def _addmath_calculus(track, topic, diff, rng, skill, source):
    families={"Foundation":["differentiate","integrate"],"Similar":["stationary","definite"],"Stretch":["optimise","kinematics"]}
    fam=rng.choice(families.get(diff,families["Similar"])); x=sp.Symbol("x")
    if fam=="differentiate":
        a,n,b=rng.randint(2,5),rng.randint(2,4),rng.randint(-5,5); expr=a*x**n+b*x; ans=sp.diff(expr,x)
        return _q(track,topic,diff,f"Differentiate y = {sp.sstr(expr)} with respect to x.",skill,["Apply the power rule."],str(ans),[f"dy/dx={ans}"],kind="expr",family=fam,source=source)
    if fam=="integrate":
        a,n=rng.randint(2,5),rng.randint(1,3); expr=a*x**n; ans=sp.integrate(expr,x)
        return _q(track,topic,diff,f"Find the indefinite integral of {sp.sstr(expr)} with respect to x.",skill,["Increase the power by one and divide by the new power."],str(ans),[f"{ans} + C"],kind="text",family=fam,source=source)
    if fam=="stationary":
        h,k=rng.randint(-4,4),rng.randint(-5,5); expr=sp.expand((x-h)**2+k)
        return _q(track,topic,diff,f"Find the stationary point of y = {expr}.",skill,["Set dy/dx=0."],f"({h},{k})",[f"({h},{k})"],kind="text",family=fam,source=source)
    if fam=="definite":
        a=rng.randint(1,4); upper=rng.randint(2,5); ans=sp.integrate(a*x,(x,0,upper))
        return _q(track,topic,diff,f"Evaluate the definite integral of {a}x from 0 to {upper}.",skill,["Integrate then apply the limits."],str(ans),[str(ans)],family=fam,source=source)
    if fam=="kinematics":
        t=sp.Symbol("t"); a,b=rng.randint(1,4),rng.randint(1,6); s=a*t**3+b*t**2; v=sp.diff(s,t)
        return _q(track,topic,diff,f"A particle has displacement s = {s}. Find velocity v as a function of t.",skill,["Velocity is ds/dt."],str(v),[f"v={v}"],kind="expr",family=fam,source=source)
    a,b=rng.randint(1,4),rng.randint(2,8); expr=-a*x**2+b*x; xv=sp.Rational(b,2*a); yv=sp.simplify(expr.subs(x,xv))
    return _q(track,topic,diff,f"Find the maximum value of y = {expr}.",skill,["Find the stationary point."],str(yv),[f"x={xv}",f"maximum={yv}"],kind="text",family=fam,source=source)


def _generator_for(topic: Topic) -> Callable:
    code = topic.code
    name = topic.name.lower()
    if code.startswith("C"): return _addmath_calculus
    if code.startswith("T") and topic.strand in {"Geometry and Trigonometry", "Calculus"}: return _addmath_trig
    if code.startswith("A") and topic.strand == "Algebra": return _addmath_algebra
    if any(k in name for k in ["four operations", "numbers", "indices", "standard form"]): return _numbers
    if "percentage" in name or "financial" in name or "practical situations" in name: return _percentage
    if re.search(r"\bratio\b|\bproportion\b|\bmap scale\b|\bscale\b", name): return _ratio
    if re.search(r"\brate\b|\bspeed\b", name): return _rate
    if any(k in name for k in ["algebraic", "factorisation", "formula", "surds", "polynomial", "partial fraction", "binomial", "logarithmic", "exponential"]): return _algebra if topic.strand != "Algebra" else _addmath_algebra
    if any(k in name for k in ["equation", "inequalit", "simultaneous", "fractional equation", "quadratic equation"]): return _equations if topic.strand != "Algebra" else _addmath_algebra
    if any(k in name for k in ["function", "graph", "coordinate geometry"]): return _functions if topic.strand != "Algebra" else _addmath_algebra
    if "sequence" in name or "nth term" in name: return _sequence
    if "circle" in name: return _circle if topic.strand != "Geometry and Trigonometry" else _addmath_trig
    if any(k in name for k in ["pythagoras", "trigonometry", "sine rule", "cosine rule", "congruence", "similarity", "arc length", "sector", "radian"]): return _pyth_trig if topic.strand != "Geometry and Trigonometry" else _addmath_trig
    if any(k in name for k in ["mensuration", "surface area", "volume", "prism", "cylinder", "pyramid", "cone", "sphere", "area and perimeter"]): return _mensuration
    if any(k in name for k in ["geometry", "angles", "triangle", "quadrilateral", "symmetry"]): return _geometry
    if any(k in name for k in ["statistics", "data", "histogram", "cumulative frequency", "measures of spread"]): return _statistics
    if "probability" in name: return _probability
    if "matrices" in name: return _matrices
    if "vector" in name: return _vectors
    if any(k in name for k in ["differentiat", "integrat", "tangent", "normal", "stationary", "maxima", "minima", "kinematic", "derivative"]): return _addmath_calculus
    return _numbers


def generate_question(track: str, topic_code: str, difficulty: str = "Similar", *, seed: int | None = None) -> Question:
    rng=random.Random(seed)
    topic=_topic(track,topic_code)
    lo,source=_learning_outcome(topic,track,rng)
    # difficulty-specific phrasing of the focus, while preserving the source-derived outcome.
    prefix={"Foundation":"Build fluency: ","Similar":"Apply the learning outcome: ","Stretch":"Reason and connect: "}.get(difficulty,"")
    skill=prefix+lo
    gen=_generator_for(topic)
    return gen(track,topic,difficulty,rng,skill,source)


def generate_similar(question: Question, *, seed: int | None = None, difficulty: str | None = None) -> Question:
    # Deliberately regenerate from the whole topic family bank and avoid repeating
    # the exact same family where the topic has alternatives.
    target_difficulty = difficulty or question.difficulty
    base_seed = int(seed or 0)
    candidate = generate_question(question.track, question.topic_code, target_difficulty, seed=base_seed)
    for offset in range(1, 8):
        if candidate.family != question.family:
            break
        candidate = generate_question(question.track, question.topic_code, target_difficulty, seed=base_seed + 104729 * offset)
    return candidate


def _last_nonempty(text: str) -> str:
    lines=[x.strip() for x in str(text or '').splitlines() if x.strip()]
    return lines[-1] if lines else ''


def _normalise_answer_text(s: str) -> str:
    s=s.strip().lower().replace('−','-').replace('×','*').replace('÷','/')
    s=s.replace('π','pi').replace('°','')
    s=re.sub(r'\b(answer|ans|x|y)\s*=\s*','',s)
    s=re.sub(r'\b(cm|km|m|kg|g|s|h|km/h|cm2|cm3|m2|m3)\b','',s)
    return s.strip(' .')


def _answer_matches(q: Question, student: str) -> bool:
    s=_normalise_answer_text(student)
    target=_normalise_answer_text(str(q.answer_display))
    if not s: return False
    if q.answer_kind in {'text'}:
        return re.sub(r'\s+','',s)==re.sub(r'\s+','',target)
    if q.answer_kind=='fraction':
        try:
            return sp.Rational(s)==sp.Rational(target)
        except Exception:
            return s==target
    if q.answer_kind=='expr':
        try:
            x=sp.symbols('x')
            return sp.simplify(sp.sympify(s,locals={'x':x})-sp.sympify(target,locals={'x':x}))==0
        except Exception:
            return re.sub(r'\s+','',s)==re.sub(r'\s+','',target)
    try:
        return abs(float(sp.N(sp.sympify(s)))-float(sp.N(sp.sympify(target))))<1e-6
    except Exception:
        # find last number in student's answer
        nums=re.findall(r'-?\d+(?:\.\d+)?',s)
        try:
            return bool(nums) and abs(float(nums[-1])-float(target))<1e-6
        except Exception:
            return s==target


def evaluate_attempt(question: Question, working_text: str) -> AttemptResult:
    lines=[x.strip() for x in str(working_text or '').splitlines() if x.strip()]
    final=_last_nonempty(working_text)
    correct=_answer_matches(question,final)
    feedback=[]
    for i,line in enumerate(lines,1):
        feedback.append(StepFeedback(i,line,"correct" if (i==len(lines) and correct) else "checked","Step recorded. Check that each transformation follows from the previous line."))
    if correct:
        return AttemptResult(True,100,90 if len(lines)>1 else 80,"Strong","Your final answer is correct.",step_feedback=feedback,strengths=["Correct final result","Relevant method shown" if len(lines)>1 else "Correct computation"],next_hint="Try a different question family from the same learning outcome.")
    return AttemptResult(False,35,50 if len(lines)>1 else 30,"Developing","The final answer does not match the verified answer.",first_logic_break=len(lines) if lines else 1,first_logic_break_explanation="Check the final computation or equation against the target skill.",step_feedback=feedback,gaps=["Recheck the final step","Use the hint to identify the governing relationship"],next_hint=question.hints[0] if question.hints else "Re-read the question and identify the required relationship.")


def analyze_own_algebra_question(question_text: str, working_text: str) -> AttemptResult:
    # Lightweight deterministic checker retained for the app's own-algebra tab.
    q=str(question_text or '')
    m=re.search(r'(?:solve\s*)?(.+?=.+?)(?:\.|$)',q,re.I)
    if not m:
        return AttemptResult(False,0,0,"Not assessed","Enter a one-variable equation such as 3(x+2)=18.",next_hint="Use an equation containing x and =.")
    x=sp.symbols('x')
    try:
        lhs,rhs=m.group(1).split('=',1)
        eq=sp.Eq(sp.sympify(lhs,locals={'x':x}),sp.sympify(rhs,locals={'x':x}))
        sols=sp.solve(eq,x)
    except Exception:
        return AttemptResult(False,0,0,"Not assessed","I could not parse the equation deterministically.",next_hint="Use * for multiplication if needed and standard algebra notation.")
    if not sols:
        target='no solution'
    else:
        target=str(sols[0])
    dummy=Question('OWN','ALG','Own algebra','Number and Algebra','Offline',q,'Solve a one-variable equation.',[],target,[f'x={target}'],'numeric',target)
    return evaluate_attempt(dummy,working_text)
