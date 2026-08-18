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


TOPICS = {
    "NT": [
        Topic("N1", "Numbers and operations", "Number and Algebra", ("integer", "negative", "operations", "numbers", "fraction", "decimal")),
        Topic("N2", "Ratio and proportion", "Number and Algebra", ("ratio", "proportion", "relative comparison")),
        Topic("N3", "Percentages", "Number and Algebra", ("percentage", "discount", "GST", "percent")),
        Topic("N4", "Rate and speed", "Number and Algebra", ("rate", "speed", "average speed", "unit")),
        Topic("N5", "Algebraic expressions and formulae", "Number and Algebra", ("algebraic expressions", "simplifying", "variable", "formula")),
        Topic("N6", "Linear equations", "Number and Algebra", ("linear equation", "equivalent equations", "solving an equation")),
        Topic("G1", "Angles and geometrical properties", "Geometry and Measurement", ("angle", "geometry", "parallel", "symmetry")),
        Topic("G2", "Mensuration", "Geometry and Measurement", ("area", "perimeter", "volume", "surface area")),
        Topic("S1", "Data handling", "Statistics and Probability", ("data", "mean", "median", "chart", "diagram")),
        Topic("P1", "Probability", "Statistics and Probability", ("probability", "chance", "event")),
    ],
    "NA": [
        Topic("N1", "Numbers and operations", "Number and Algebra", ("numbers", "indices", "significant figures", "rounding")),
        Topic("N2", "Ratio, rate and proportion", "Number and Algebra", ("ratio", "rate", "proportion", "scale")),
        Topic("N3", "Percentages and financial contexts", "Number and Algebra", ("percentage", "discount", "interest")),
        Topic("N4", "Algebraic expressions", "Number and Algebra", ("algebraic expressions", "expand", "factorising", "identities")),
        Topic("N5", "Equations and inequalities", "Number and Algebra", ("equation", "inequality", "simultaneous")),
        Topic("N6", "Functions and graphs", "Number and Algebra", ("functions", "graph", "input", "output", "linear functions")),
        Topic("G1", "Similarity and scale", "Geometry and Measurement", ("similar", "scale", "proportional relationship")),
        Topic("G2", "Pythagoras and trigonometry", "Geometry and Measurement", ("Pythagoras", "trigonometric", "right-angled")),
        Topic("G3", "Mensuration", "Geometry and Measurement", ("area", "volume", "surface area")),
        Topic("S1", "Statistics", "Statistics and Probability", ("statistics", "mean", "median", "histogram", "spread")),
        Topic("P1", "Probability", "Statistics and Probability", ("probability", "tree diagram", "possibility")),
    ],
    "O": [
        Topic("N1", "Numbers and operations", "Number and Algebra", ("prime factorisation", "real numbers", "significant figures", "indices")),
        Topic("N2", "Ratio and proportion", "Number and Algebra", ("ratio", "proportionality", "percentage")),
        Topic("N3", "Rates and speed", "Number and Algebra", ("rate", "speed", "average speed")),
        Topic("N4", "Algebraic expressions and identities", "Number and Algebra", ("algebraic expressions", "quadratic", "identities", "factorising")),
        Topic("N5", "Equations and inequalities", "Number and Algebra", ("equation", "simultaneous", "inequality")),
        Topic("N6", "Functions and graphs", "Number and Algebra", ("function", "quadratic functions", "graph", "linear functions")),
        Topic("N7", "Sequences", "Number and Algebra", ("sequence", "general term", "geometric patterns")),
        Topic("G1", "Geometry and angles", "Geometry and Measurement", ("angle", "geometrical", "parallel", "congruency")),
        Topic("G2", "Pythagoras, similarity and trigonometry", "Geometry and Measurement", ("Pythagoras", "similar", "trigonometric")),
        Topic("G3", "Mensuration", "Geometry and Measurement", ("mensuration", "area", "volume", "surface area")),
        Topic("G4", "Circle properties", "Geometry and Measurement", ("circle", "chord", "tangent")),
        Topic("S1", "Statistics", "Statistics and Probability", ("data", "histogram", "standard deviation", "interquartile")),
        Topic("P1", "Probability", "Statistics and Probability", ("probability", "tree diagram", "possibility")),
        Topic("N8", "Matrices", "Number and Algebra", ("matrices", "matrix")),
        Topic("G5", "Vectors", "Geometry and Measurement", ("vectors", "magnitude", "direction")),
    ],
}


def topics_for_track(track: str) -> list[Topic]:
    return list(TOPICS.get(track, TOPICS["O"]))


def official_topic_code(track: str, code: str) -> str:
    return code



OUTCOME_FOCI = {
    "NT": {
        "N1": [
            "Use positive and negative numbers in real-life contexts and perform the four operations accurately.",
            "Estimate quantities and judge whether numerical answers are reasonable.",
            "Connect fractions, decimals and numbers represented in different forms.",
        ],
        "N2": [
            "Represent and interpret ratios and proportions in real-life situations.",
            "Generate equivalent ratios and connect ratios with fractions and decimals.",
            "Use proportional reasoning to solve comparison and sharing problems.",
        ],
        "N3": [
            "Use percentages in everyday contexts such as discounts, charges and money problems.",
            "Connect percentages with fractions and decimals.",
            "Find percentage change and reverse percentages in practical situations.",
        ],
        "N4": [
            "Understand rate as a comparison of two quantities with derived units.",
            "Distinguish speed from average speed and solve rate problems in context.",
            "Use direct or inverse proportional reasoning in rate situations.",
        ],
        "N5": [
            "Interpret and simplify algebraic expressions with integral coefficients.",
            "Use algebraic notation to represent relationships between quantities.",
            "Substitute values into algebraic expressions and formulae accurately.",
        ],
        "N6": [
            "Understand equations as statements of equality and solve linear equations by maintaining equivalence.",
            "Form linear equations from everyday situations and solve them algebraically.",
            "Check whether a value is a solution of a linear equation.",
        ],
        "G1": [
            "Identify, classify and calculate angles using standard geometrical properties.",
            "Use geometrical notation and spatial reasoning to interpret diagrams.",
            "Recognise symmetry and relationships between lines and angles.",
        ],
        "G2": [
            "Use perimeter, area, volume and surface area as measures of boundaries and space.",
            "Model real objects using basic 2D shapes and 3D solids.",
            "Compose and decompose shapes while conserving area or volume.",
        ],
        "S1": [
            "Organise, represent, analyse and interpret data in appropriate statistical representations.",
            "Calculate and interpret common measures of centre and spread.",
            "Reason about how data values affect summary statistics.",
        ],
        "P1": [
            "Interpret probability as a measure of chance between 0 and 1.",
            "Calculate probabilities from equally likely outcomes.",
            "Use complements and simple multi-stage probability reasoning.",
        ],
    },
    "NA": {
        "N1": [
            "Use real numbers, indices, rounding and significant figures accurately.",
            "Estimate and check the reasonableness of numerical results.",
            "Use equivalent numerical forms strategically in calculations.",
        ],
        "N2": [
            "Interpret ratios, rates, scales and direct or inverse proportion using tables, equations and contexts.",
            "Use scale drawings and maps to determine actual lengths and areas.",
            "Solve proportional comparison and sharing problems.",
        ],
        "N3": [
            "Apply percentages in financial and real-world contexts.",
            "Find percentage increase, decrease and reverse percentages.",
            "Connect percentage calculations with proportional reasoning.",
        ],
        "N4": [
            "Expand and factorise algebraic expressions and recognise these as reverse processes.",
            "Use special algebraic identities correctly.",
            "Substitute into and manipulate algebraic formulae.",
        ],
        "N5": [
            "Form and solve linear equations, simultaneous equations and inequalities.",
            "Connect algebraic solutions with graphical representations.",
            "Interpret inequality notation and represent solution sets appropriately.",
        ],
        "N6": [
            "Represent functions using input-output rules, tables, equations and coordinate graphs.",
            "Interpret gradients and intercepts in linear functions.",
            "Use functions to model simple real-world relationships.",
        ],
        "G1": [
            "Use similarity and scale factors to compare corresponding lengths and areas.",
            "Apply proportional reasoning to similar figures and scale drawings.",
            "Recognise and justify similarity from geometrical information.",
        ],
        "G2": [
            "Apply Pythagoras' theorem and trigonometric ratios in right-angled triangles.",
            "Use diagrams and notation to communicate geometrical reasoning.",
            "Apply right-triangle methods in practical or unfamiliar contexts.",
        ],
        "G3": [
            "Calculate and connect area, surface area and volume of related shapes and solids.",
            "Model practical objects using standard geometrical solids.",
            "Use composition and decomposition to solve mensuration problems.",
        ],
        "S1": [
            "Construct, interpret and compare statistical representations and summary measures.",
            "Use measures of centre and spread to compare data sets.",
            "Reason about the effect of extreme values on statistical measures.",
        ],
        "P1": [
            "Use probability notation and calculate probabilities of events.",
            "Organise multi-stage outcomes using systematic representations.",
            "Use complement and conditional reasoning in simple probability problems.",
        ],
    },
    "O": {
        "N1": [
            "Use prime factorisation, real numbers, indices, significant figures and rounding appropriately.",
            "Reason about equivalent number representations and numerical invariance.",
            "Estimate and evaluate the effect of rounding on calculations.",
        ],
        "N2": [
            "Use ratios, fractions, decimals and percentages as equivalent proportional representations.",
            "Solve proportional sharing and comparison problems.",
            "Use reverse proportional reasoning in unfamiliar contexts.",
        ],
        "N3": [
            "Interpret rate as a compound measure and distinguish speed from average speed.",
            "Solve multi-stage rate and average-speed problems.",
            "Use units consistently when reasoning about rates.",
        ],
        "N4": [
            "Manipulate algebraic expressions, including expansion, factorisation and identities.",
            "Recognise equivalent algebraic forms and select useful forms for a purpose.",
            "Substitute into and rearrange algebraic expressions and formulae.",
        ],
        "N5": [
            "Solve equations, simultaneous equations and inequalities and connect them with graphical meaning.",
            "Form equations from contextual information and justify algebraic steps.",
            "Interpret solution sets and verify solutions.",
        ],
        "N6": [
            "Represent and interpret functions algebraically, numerically and graphically.",
            "Use linear and quadratic functions to model relationships.",
            "Connect features of an equation with features of its graph.",
        ],
        "N7": [
            "Describe number patterns and formulate general terms of sequences.",
            "Use equivalent expressions to represent the nth term of a sequence.",
            "Work backwards from a term value or rule to determine position in a sequence.",
        ],
        "G1": [
            "Use geometrical properties, notation and logical reasoning to calculate angles.",
            "Interpret geometrical diagrams and justify relationships between lines and angles.",
            "Apply angle properties in multi-step configurations.",
        ],
        "G2": [
            "Apply Pythagoras' theorem, similarity and trigonometric ratios to solve geometrical problems.",
            "Use proportional relationships in similar figures.",
            "Interpret diagrams and choose an efficient right-triangle method.",
        ],
        "G3": [
            "Solve mensuration problems involving area, surface area and volume.",
            "Compose and decompose shapes and solids to calculate measures.",
            "Model practical objects using related geometrical shapes and solids.",
        ],
        "G4": [
            "Investigate and apply invariant circle properties involving chords, angles and tangents.",
            "Use correct circle terminology and geometrical reasoning.",
            "Apply circle properties in multi-step problems.",
        ],
        "S1": [
            "Interpret statistical diagrams and use measures of centre and spread to compare data sets.",
            "Reason about range, interquartile range and standard deviation.",
            "Analyse how individual or extreme values affect statistical measures.",
        ],
        "P1": [
            "Organise outcomes systematically and calculate probabilities of single and combined events.",
            "Use tree or possibility reasoning for multi-stage probability.",
            "Apply complement and conditional probability reasoning.",
        ],
        "N8": [
            "Use matrices to represent information and perform matrix operations using correct notation.",
            "Interpret matrix calculations in practical contexts.",
            "Communicate matrix relationships accurately using standard notation.",
        ],
        "G5": [
            "Represent vectors by magnitude and direction and use vector notation correctly.",
            "Add, subtract and resolve vectors using components.",
            "Use vectors to describe displacement and relationships between points.",
        ],
    },
}

def _level_for_track(track: str) -> str:
    return {"NT": "G1", "NA": "G2", "O": "G3"}.get(track, "G3")


def _learning_outcome(topic: Topic, track: str, rng: random.Random) -> tuple[str, str]:
    focuses = OUTCOME_FOCI.get(track, {}).get(topic.code, [])
    if focuses:
        focus = rng.choice(focuses)
        level = _level_for_track(track)
        # Keep a traceable source section from the uploaded workbook.
        sections = [k for k in LEARNING_OUTCOMES if k.endswith(level)]
        source = rng.choice(sections) if sections else f"{level} learning outcomes"
        return focus, source
    return (f"Apply {topic.name.lower()} accurately in an appropriate problem.", f"{_level_for_track(track)} learning outcomes")


def _topic(track: str, code: str) -> Topic:
    for t in topics_for_track(track):
        if t.code == code:
            return t
    return topics_for_track(track)[0]


def _q(track, topic, difficulty, prompt, skill, hints, answer, solution, kind="numeric", family="", source=""):
    return Question(track, topic.code, topic.name, topic.strand, difficulty, prompt, skill, hints, answer, solution, kind, answer, family, source)


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


def _statistics(track,topic,diff,rng,skill,source):
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


def _matrices(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["add","multiply_scalar"],"Similar":["add","product"],"Stretch":["product"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    a,b,c,d=[rng.randint(-5,8) for _ in range(4)]
    if fam=="add":
        e,f,g,h=[rng.randint(-5,8) for _ in range(4)]; ans=[[a+e,b+f],[c+g,d+h]]
        prompt=f"Given A = [[{a},{b}],[{c},{d}]] and B = [[{e},{f}],[{g},{h}]], find A + B."
    elif fam=="multiply_scalar":
        k=rng.randint(2,5); ans=[[k*a,k*b],[k*c,k*d]]; prompt=f"Given A = [[{a},{b}],[{c},{d}]], find {k}A."
    else:
        # matrix by column vector
        x,y=rng.randint(1,6),rng.randint(1,6); ans=[a*x+b*y,c*x+d*y]; prompt=f"Calculate [[{a},{b}],[{c},{d}]] [[{x}],[{y}]]."
    disp=str(ans)
    return _q(track,topic,diff,prompt,skill,["Use the appropriate matrix operation row by row."],disp,[f"Answer = {disp}"],kind="text",family=fam,source=source)


def _vectors(track,topic,diff,rng,skill,source):
    family_bank={"Foundation":["component","magnitude"],"Similar":["component","position"],"Stretch":["position","magnitude"]}
    fam=rng.choice(family_bank.get(diff,family_bank["Similar"]))
    a,b=rng.randint(-6,8),rng.randint(-6,8)
    if fam=="component":
        c,d=rng.randint(-6,8),rng.randint(-6,8); ans=f"({a+c},{b+d})"; prompt=f"Given vector u = ({a},{b}) and v = ({c},{d}), find u + v."
    elif fam=="magnitude":
        # use a Pythagorean pair
        a,b=rng.choice([(3,4),(5,12),(8,15)]); ans=str(int(math.hypot(a,b))); prompt=f"Find the magnitude of vector ({a},{b})."
    else:
        c,d=rng.randint(-5,8),rng.randint(-5,8); ans=f"({c-a},{d-b})"; prompt=f"Point A is ({a},{b}) and point B is ({c},{d}). Find vector AB."
    return _q(track,topic,diff,prompt,skill,["Use vector components consistently."],ans,[f"Answer = {ans}"],kind="text",family=fam,source=source)


def _generator_for(topic: Topic) -> Callable:
    code = topic.code
    name = topic.name.lower()
    if code == "N1": return _numbers
    if "percentage" in name or "financial" in name: return _percentage
    if re.search(r"\bratio\b|\bproportion\b|\bscale\b", name): return _ratio
    if re.search(r"\brate\b|\bspeed\b", name): return _rate
    if "algebraic" in name: return _algebra
    if "equations" in name or "inequalities" in name: return _equations
    if "functions" in name or "graphs" in name: return _functions
    if "sequence" in name: return _sequence
    if "circle" in name: return _circle
    if "pythagoras" in name or "trigonometry" in name or "similarity" in name: return _pyth_trig
    if "mensuration" in name: return _mensuration
    if "geometry" in name or "angles" in name: return _geometry
    if "statistics" in name or "data" in name: return _statistics
    if "probability" in name: return _probability
    if "matrices" in name: return _matrices
    if "vectors" in name: return _vectors
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
