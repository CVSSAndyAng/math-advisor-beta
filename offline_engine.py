from __future__ import annotations

import math
import random
import re
import statistics
from dataclasses import dataclass, field
from fractions import Fraction
from typing import Any, Callable

import sympy as sp
from sympy.parsing.sympy_parser import (
    implicit_multiplication_application,
    standard_transformations,
    convert_xor,
    parse_expr,
)

TRANSFORMS = standard_transformations + (convert_xor, implicit_multiplication_application)

TRACKS = {
    "O-Level Mathematics (4052)": "O",
    "N(A)-Level Mathematics Syllabus A (4045)": "NA",
    "N(T)-Level Mathematics Syllabus T (4046)": "NT",
}


@dataclass(frozen=True)
class Topic:
    code: str
    strand: str
    name: str
    tracks: tuple[str, ...]
    offline_support: str
    notes: str


TOPICS: tuple[Topic, ...] = (
    Topic("N1", "Number and Algebra", "Numbers and operations", ("O", "NA", "NT"), "Strong", "Arithmetic, HCF/LCM where applicable, rounding, standard form, indices."),
    Topic("N2", "Number and Algebra", "Ratio and proportion", ("O", "NA", "NT"), "Strong", "Ratios, scales, direct and inverse proportion."),
    Topic("N3", "Number and Algebra", "Percentage and finance", ("O", "NA", "NT"), "Strong", "Percentage change, reverse percentage, simple/compound-interest style calculations."),
    Topic("N4", "Number and Algebra", "Rate and speed", ("O", "NA", "NT"), "Strong", "Average speed/rate and unit conversion."),
    Topic("N5", "Number and Algebra", "Algebraic expressions and formulae", ("O", "NA", "NT"), "Strong", "Expansion, factorisation, substitution, formulae and sequences."),
    Topic("N6", "Number and Algebra", "Functions and graphs", ("O", "NA", "NT"), "Partial", "Generated coordinate/gradient/function-value questions; freehand graph interpretation needs human/AI vision."),
    Topic("N7", "Number and Algebra", "Equations and inequalities", ("O", "NA", "NT"), "Strong", "Linear, simultaneous and quadratic equations; inequalities for O/NA."),
    Topic("N8", "Number and Algebra", "Set language and notation", ("O",), "Strong", "Union, intersection, complement and element counting in generated problems."),
    Topic("N9", "Number and Algebra", "Matrices", ("O",), "Strong", "Matrix addition, scalar multiplication and multiplication."),
    Topic("G1", "Geometry and Measurement", "Angles, triangles and polygons", ("O", "NA", "NT"), "Strong", "Angle facts and polygon angle sums in generated diagram-free problems."),
    Topic("G2", "Geometry and Measurement", "Congruence, similarity and symmetry", ("O", "NA", "NT"), "Strong", "Scale factors, corresponding lengths, area/volume scale factors where applicable."),
    Topic("G3", "Geometry and Measurement", "Circle properties", ("O", "NA"), "Partial", "Generated circle-theorem problems without diagrams; arbitrary uploaded diagrams are not read offline."),
    Topic("G4", "Geometry and Measurement", "Pythagoras and trigonometry", ("O", "NA", "NT"), "Strong", "Right-triangle trigonometry for all tracks; sine/cosine rule and bearings for O/NA."),
    Topic("G5", "Geometry and Measurement", "Mensuration", ("O", "NA", "NT"), "Strong", "Perimeter, area, volume, surface area, sectors and composite-style calculations."),
    Topic("G6", "Geometry and Measurement", "Coordinate geometry", ("O", "NA"), "Strong", "Gradient, distance and line equations."),
    Topic("G7", "Geometry and Measurement", "Vectors in two dimensions", ("O",), "Strong", "Vector arithmetic, magnitudes and simple position-vector problems."),
    Topic("S1", "Statistics and Probability", "Data handling and analysis", ("O", "NA", "NT"), "Strong", "Mean, median, mode, quartiles, range; standard deviation for O/NA."),
    Topic("S2", "Statistics and Probability", "Probability", ("O", "NA", "NT"), "Strong", "Single and combined events; independent/mutually exclusive events for O/NA."),
)


@dataclass
class Question:
    id: str
    track: str
    topic_code: str
    topic_name: str
    strand: str
    difficulty: str
    prompt: str
    target_skill: str
    hints: list[str]
    worked_solution: list[str]
    answer_display: str
    checker: str
    expected: Any
    tolerance: float = 1e-6
    metadata: dict[str, Any] = field(default_factory=dict)


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
    summary: str
    first_logic_break: int | None
    first_logic_break_explanation: str
    step_feedback: list[StepFeedback]
    strengths: list[str]
    gaps: list[str]
    next_hint: str
    mastery: str


def topics_for_track(track_code: str) -> list[Topic]:
    return [t for t in TOPICS if track_code in t.tracks]


def official_topic_code(track_code: str, internal_code: str) -> str:
    # N(T) uses G3 for Pythagoras/trigonometry and G4 for mensuration,
    # whereas the O/N(A) syllabuses use G4 and G5 respectively.
    if track_code == "NT" and internal_code == "G4":
        return "G3"
    if track_code == "NT" and internal_code == "G5":
        return "G4"
    return internal_code


def _parse_expr(text: str) -> sp.Expr:
    cleaned = text.strip().replace("×", "*").replace("÷", "/").replace("−", "-")
    cleaned = re.sub(r"(?<=\d),(?=\d{3}(?:\D|$))", "", cleaned)
    return parse_expr(cleaned, transformations=TRANSFORMS, evaluate=True)


def _parse_equation(text: str) -> sp.Eq:
    cleaned = text.strip().replace("×", "*").replace("÷", "/").replace("−", "-")
    if cleaned.lower().startswith("solve"):
        cleaned = re.sub(r"^solve\s*", "", cleaned, flags=re.I)
    cleaned = cleaned.rstrip(". ")
    if "=" not in cleaned:
        raise ValueError("No equals sign")
    left, right = cleaned.split("=", 1)
    return sp.Eq(_parse_expr(left), _parse_expr(right), evaluate=False)


def _equation_solution_signature(eq: sp.Eq) -> tuple[str, ...] | None:
    symbols = sorted(eq.free_symbols, key=lambda s: s.name)
    if len(symbols) != 1:
        return None
    x = symbols[0]
    try:
        sols = sp.solve(eq, x)
    except Exception:
        return None
    norm = []
    for sol in sols:
        try:
            norm.append(sp.srepr(sp.simplify(sol)))
        except Exception:
            norm.append(str(sol))
    return tuple(sorted(norm))


def equations_equivalent(a: sp.Eq, b: sp.Eq) -> bool:
    sig_a = _equation_solution_signature(a)
    sig_b = _equation_solution_signature(b)
    if sig_a is not None and sig_b is not None:
        return sig_a == sig_b
    # Fallback: lhs-rhs expressions are non-zero constant multiples.
    da = sp.expand(a.lhs - a.rhs)
    db = sp.expand(b.lhs - b.rhs)
    if sp.simplify(da - db) == 0 or sp.simplify(da + db) == 0:
        return True
    try:
        ratio = sp.simplify(da / db)
        return bool(ratio.free_symbols == set() and ratio != 0)
    except Exception:
        return False


def _last_nonempty_line(text: str) -> str:
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    return lines[-1] if lines else text.strip()


def _extract_number(text: str) -> float | None:
    # Prefer the RHS of the final equality if present.
    candidate = _last_nonempty_line(text)
    if "=" in candidate:
        candidate = candidate.split("=")[-1]
    candidate = candidate.strip().replace("°", "").replace("%", "")
    candidate = re.sub(r"\([^)]*\)\s*$", "", candidate).strip()
    nums = re.findall(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?(?:/[0-9]+)?", candidate)
    if not nums:
        return None
    raw = nums[-1]
    try:
        if "/" in raw:
            return float(Fraction(raw))
        return float(raw)
    except Exception:
        return None


def _normalize_ratio(text: str) -> tuple[int, ...] | None:
    candidate = _last_nonempty_line(text)
    if "=" in candidate:
        candidate = candidate.split("=")[-1]
    m = re.search(r"(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)", candidate)
    if not m:
        return None
    a = Fraction(m.group(1))
    b = Fraction(m.group(2))
    if b == 0:
        return None
    ratio = a / b
    # Canonical integer pair.
    return (ratio.numerator, ratio.denominator)


def _parse_pair(text: str) -> tuple[float, float] | None:
    # Recognise common labelled pairs first.
    h = re.findall(r"HCF[^=]*=\s*([-+]?\d+(?:\.\d+)?)", text, re.I)
    l = re.findall(r"LCM[^=]*=\s*([-+]?\d+(?:\.\d+)?)", text, re.I)
    if h and l:
        return float(h[-1]), float(l[-1])
    xs = re.findall(r"\bx\s*=\s*([-+]?\d+(?:\.\d+)?)", text, re.I)
    ys = re.findall(r"\by\s*=\s*([-+]?\d+(?:\.\d+)?)", text, re.I)
    if xs and ys:
        return float(xs[-1]), float(ys[-1])

    candidate = _last_nonempty_line(text)
    nums = re.findall(r"[-+]?\d+(?:\.\d+)?", candidate)
    if len(nums) >= 2:
        return float(nums[-2]), float(nums[-1])
    return None


def _parse_roots(text: str) -> set[float] | None:
    nums = re.findall(r"[-+]?\d+(?:\.\d+)?(?:/[0-9]+)?", text)
    if not nums:
        return None
    vals: set[float] = set()
    for raw in nums[-4:]:
        try:
            vals.add(float(Fraction(raw)) if "/" in raw else float(raw))
        except Exception:
            pass
    return vals or None


def _parse_matrix(text: str) -> list[list[float]] | None:
    rows = re.findall(r"\[\s*([^\[\]]+?)\s*\]", text)
    if not rows:
        # Accept 1,2;3,4
        if ";" in text:
            rows = text.strip().strip("[]()").split(";")
        else:
            return None
    matrix: list[list[float]] = []
    for row in rows:
        nums = re.findall(r"[-+]?\d+(?:\.\d+)?", row)
        if nums:
            matrix.append([float(x) for x in nums])
    return matrix or None


def _answer_correct(question: Question, attempt_text: str) -> bool:
    checker = question.checker
    expected = question.expected
    if checker == "numeric":
        got = _extract_number(attempt_text)
        return got is not None and math.isclose(got, float(expected), rel_tol=1e-6, abs_tol=question.tolerance)
    if checker == "numeric_percent":
        got = _extract_number(attempt_text)
        return got is not None and math.isclose(got, float(expected), rel_tol=1e-6, abs_tol=question.tolerance)
    if checker == "ratio":
        return _normalize_ratio(attempt_text) == tuple(expected)
    if checker == "expression":
        candidate = _last_nonempty_line(attempt_text)
        if "=" in candidate:
            candidate = candidate.split("=", 1)[1]
        try:
            return sp.simplify(_parse_expr(candidate) - _parse_expr(str(expected))) == 0
        except Exception:
            return False
    if checker == "equation_solution":
        got = _extract_number(attempt_text)
        return got is not None and math.isclose(got, float(expected), rel_tol=1e-6, abs_tol=question.tolerance)
    if checker == "quadratic_roots":
        got = _parse_roots(attempt_text)
        exp = {float(x) for x in expected}
        return got is not None and exp.issubset(got) and len(got) <= max(2, len(exp) + 1)
    if checker == "pair":
        got = _parse_pair(attempt_text)
        if got is None:
            # Try x=..., y=...
            xs = re.findall(r"x\s*=\s*([-+]?\d+(?:\.\d+)?)", attempt_text, re.I)
            ys = re.findall(r"y\s*=\s*([-+]?\d+(?:\.\d+)?)", attempt_text, re.I)
            if xs and ys:
                got = (float(xs[-1]), float(ys[-1]))
        return got is not None and all(math.isclose(got[i], float(expected[i]), abs_tol=question.tolerance) for i in (0, 1))
    if checker == "matrix":
        got = _parse_matrix(attempt_text)
        if got is None:
            return False
        exp = [[float(v) for v in row] for row in expected]
        return got == exp
    if checker == "set":
        candidate = _last_nonempty_line(attempt_text)
        brace = re.findall(r"\{([^{}]*)\}", candidate)
        if brace:
            candidate = brace[-1]
        vals = set(re.findall(r"[-+]?\d+(?:\.\d+)?|\b[A-Za-z]\b", candidate))
        return vals == set(str(x) for x in expected)
    if checker == "inequality":
        candidate = _last_nonempty_line(attempt_text).replace("≤", "<=").replace("≥", ">=")
        expected_s = str(expected).replace("<=", "≤").replace(">=", "≥")
        norm = candidate.replace(" ", "").replace("≤", "<=").replace("≥", ">=")
        exp_norm = str(expected).replace(" ", "")
        return exp_norm in norm or norm.endswith(exp_norm)
    return False


def _line_reasoning(question: Question, attempt_text: str) -> tuple[list[StepFeedback], int | None, str, int]:
    lines = [x.strip() for x in attempt_text.splitlines() if x.strip()]
    if not lines:
        return [], None, "No working was submitted.", 0

    feedback: list[StepFeedback] = []
    first_break: int | None = None
    first_expl = "No invalid algebraic transition was detected in the parseable working."

    # Algebra equation questions get true equivalence checking.
    if question.checker in {"equation_solution", "quadratic_roots", "pair", "inequality"} and question.metadata.get("equations"):
        reference_eq_text = question.metadata["equations"][0]
        try:
            reference_eq = _parse_equation(reference_eq_text)
        except Exception:
            reference_eq = None

        parseable = 0
        correct_steps = 0
        for idx, line in enumerate(lines, 1):
            # Simultaneous equations and inequalities use lighter checks.
            if question.checker in {"pair", "inequality"}:
                feedback.append(StepFeedback(idx, line, "checked", "Recorded as part of the submitted method; the final result is checked exactly."))
                continue
            try:
                eq = _parse_equation(line)
                parseable += 1
                if reference_eq is not None and equations_equivalent(reference_eq, eq):
                    correct_steps += 1
                    feedback.append(StepFeedback(idx, line, "correct", "This equation is equivalent to the original equation."))
                else:
                    if first_break is None:
                        first_break = idx
                        first_expl = "This is the first parseable line that is not equivalent to the original equation. Check the operation used just before this line."
                    feedback.append(StepFeedback(idx, line, "incorrect", "This line changes the solution set, so the algebraic transformation is not equivalent."))
            except Exception:
                feedback.append(StepFeedback(idx, line, "unparsed", "I could not reliably parse this line offline. Keep each algebraic line in a form such as 3x + 2 = 11."))
        reasoning_score = int(round(100 * correct_steps / parseable)) if parseable else 55
        return feedback, first_break, first_expl, reasoning_score

    # For other topics, validate self-contained numeric equalities where possible.
    evaluated = 0
    valid = 0
    for idx, line in enumerate(lines, 1):
        if "=" in line and line.count("=") == 1:
            left, right = line.split("=", 1)
            try:
                lv = float(sp.N(_parse_expr(left)))
                rv = float(sp.N(_parse_expr(right.replace("°", ""))))
                evaluated += 1
                if math.isclose(lv, rv, rel_tol=1e-6, abs_tol=1e-6):
                    valid += 1
                    feedback.append(StepFeedback(idx, line, "correct", "This numerical equality is valid."))
                else:
                    if first_break is None:
                        first_break = idx
                        first_expl = "This is the first numerical equality that does not balance. Recheck the calculation or formula substitution on this line."
                    feedback.append(StepFeedback(idx, line, "incorrect", "The two sides are not numerically equal."))
            except Exception:
                feedback.append(StepFeedback(idx, line, "checked", "This step could not be fully parsed, but it is retained as evidence of your method."))
        else:
            feedback.append(StepFeedback(idx, line, "checked", "This step is retained as evidence of your method; offline semantic checking is limited for prose."))
    reasoning_score = int(round(100 * valid / evaluated)) if evaluated else 70
    return feedback, first_break, first_expl, reasoning_score


def evaluate_attempt(question: Question, attempt_text: str) -> AttemptResult:
    correct = _answer_correct(question, attempt_text)
    step_feedback, first_break, first_expl, reasoning_score = _line_reasoning(question, attempt_text)
    answer_score = 100 if correct else 0
    if correct and reasoning_score >= 85:
        mastery = "Secure"
    elif correct:
        mastery = "Developing"
    elif reasoning_score >= 70:
        mastery = "Developing"
    else:
        mastery = "Emerging"

    strengths: list[str] = []
    gaps: list[str] = []
    if correct:
        strengths.append("The final answer matches the verified answer for this generated question.")
    else:
        gaps.append("The final answer does not yet match the verified answer.")
    if first_break is None and step_feedback:
        strengths.append("No invalid parseable algebraic/numerical step was detected before the final answer.")
    elif first_break is not None:
        gaps.append(f"Revisit line {first_break}; it is the earliest detected logic break.")

    if not attempt_text.strip():
        summary = "Enter your working so the tutor can check both the answer and the method."
    elif correct and first_break is None:
        summary = "Correct. Your submitted work reaches the verified answer without a detected logic break in the parts the offline checker can parse."
    elif correct:
        summary = "Your final answer is correct, but the checker found an earlier step that is not equivalent or numerically valid."
    else:
        summary = "The answer needs correction. Use the first logic break and the next hint to repair the method rather than restarting blindly."

    next_hint = question.hints[0] if question.hints else "Compare your method with the target skill."
    if first_break is not None and len(question.hints) > 1:
        next_hint = question.hints[1]

    return AttemptResult(
        is_correct=correct,
        answer_score=answer_score,
        reasoning_score=reasoning_score,
        summary=summary,
        first_logic_break=first_break,
        first_logic_break_explanation=first_expl,
        step_feedback=step_feedback,
        strengths=strengths,
        gaps=gaps,
        next_hint=next_hint,
        mastery=mastery,
    )


def _qid(rng: random.Random, topic: str) -> str:
    return f"{topic.lower()}-{rng.randint(100000, 999999)}"


def _q(track: str, topic: Topic, difficulty: str, prompt: str, skill: str, hints: list[str], solution: list[str], answer: str, checker: str, expected: Any, *, tol: float = 1e-6, metadata: dict[str, Any] | None = None, rng: random.Random) -> Question:
    return Question(_qid(rng, topic.code), track, topic.code, topic.name, topic.strand, difficulty, prompt, skill, hints, solution, answer, checker, expected, tol, metadata or {})


def generate_question(track: str, topic_code: str, difficulty: str = "Similar", seed: int | None = None) -> Question:
    rng = random.Random(seed)
    topic = next(t for t in TOPICS if t.code == topic_code and track in t.tracks)
    fn = GENERATORS[topic_code]
    return fn(track, topic, difficulty, rng)


def generate_similar(question: Question, seed: int | None = None, difficulty: str | None = None) -> Question:
    return generate_question(question.track, question.topic_code, difficulty or question.difficulty, seed)


def _gen_n1(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    variant = rng.choice(["standard_form", "rounding", "indices"] if track == "NT" else ["hcf_lcm", "standard_form", "rounding", "indices"])
    if variant == "hcf_lcm":
        a = rng.choice([24, 36, 42, 48, 60, 72]); b = rng.choice([30, 45, 54, 66, 84, 90])
        h = math.gcd(a, b); l = abs(a*b)//h
        prompt = f"Find the HCF and LCM of {a} and {b}."
        return _q(track, topic, diff, prompt, "Prime factors / common factors and multiples", ["List prime factors or common factors first.", "For HCF take common prime factors with the smaller powers; for LCM take all needed prime factors."], [f"HCF({a}, {b}) = {h}", f"LCM({a}, {b}) = {l}"], f"HCF = {h}, LCM = {l}", "pair", (h, l), rng=rng)
    if variant == "standard_form":
        mant = rng.choice([2.4, 3.6, 4.8, 7.2]); power = rng.choice([-5, -4, 4, 5, 6])
        prompt = f"Write {mant} × 10^{power} as an ordinary number."
        ans = mant * (10 ** power)
        return _q(track, topic, diff, prompt, "Convert between standard form and ordinary form", ["A positive power moves the decimal point right; a negative power moves it left.", "Move the decimal point exactly the number of places shown by the power."], [f"{mant} × 10^{power} = {ans:g}"], f"{ans:g}", "numeric", ans, tol=max(1e-10, abs(ans)*1e-9), rng=rng)
    if variant == "indices":
        base = rng.choice([2,3,5]); p = rng.randint(2,5); q = rng.randint(1,4)
        prompt = f"Simplify {base}^{p} × {base}^{q}. Give your answer as a power of {base}."
        exp = p+q
        return _q(track, topic, diff, prompt, "Use the laws of indices", ["The bases are the same.", "When multiplying powers with the same base, add the indices."], [f"{base}^{p} × {base}^{q} = {base}^{p+q}"], f"{base}^{exp}", "expression", f"{base}**{exp}", rng=rng)
    value = rng.uniform(10, 900)
    dp = rng.choice([1,2])
    ans = round(value, dp)
    prompt = f"Round {value:.4f} to {dp} decimal place{'s' if dp != 1 else ''}."
    return _q(track, topic, diff, prompt, "Rounding and approximation", ["Locate the required decimal place.", "Look at the next digit to decide whether to round up."], [f"{value:.4f} rounds to {ans:.{dp}f}."], f"{ans:.{dp}f}", "numeric", ans, tol=10**(-dp-2), rng=rng)


def _gen_n2(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    a, b = rng.choice([(2,3),(3,5),(4,7),(5,8)])
    total = (a+b)*rng.randint(6,18)
    x = total*a/(a+b); y = total*b/(a+b)
    prompt = f"Divide ${total} in the ratio {a}:{b}."
    return _q(track, topic, diff, prompt, "Divide a quantity in a given ratio", ["Add the ratio parts first.", "Find the value of one part, then multiply by each ratio number."], [f"Total parts = {a+b}", f"One part = {total} ÷ {a+b} = {total/(a+b):g}", f"Shares = ${x:g} and ${y:g}"], f"${x:g} and ${y:g}", "pair", (x,y), rng=rng)


def _gen_n3(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    if diff == "Stretch" and track in {"O","NA"}:
        original = rng.choice([80,120,160,240]); inc = rng.choice([15,20,25])
        new = original*(1+inc/100)
        prompt = f"After a {inc}% increase, a price is ${new:.2f}. Find the original price."
        return _q(track, topic, diff, prompt, "Reverse percentage", ["The new price represents more than 100% of the original.", f"Treat ${new:.2f} as {100+inc}% and divide by {(100+inc)/100:g}."], [f"Original = {new:.2f} ÷ {(100+inc)/100:g} = {original:.2f}"], f"${original:.2f}", "numeric", original, tol=0.01, rng=rng)
    original = rng.choice([60,80,120,150,240]); pct = rng.choice([10,15,20,25,30])
    new = original*(1-pct/100)
    prompt = f"A jacket costs ${original}. It is discounted by {pct}%. Find the sale price."
    return _q(track, topic, diff, prompt, "Percentage decrease", ["Find the discount or use a multiplier.", f"After a {pct}% discount, {100-pct}% remains."], [f"Sale price = {original} × {(100-pct)/100:g} = {new:.2f}"], f"${new:.2f}", "numeric", new, tol=0.01, rng=rng)


def _gen_n4(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    distance = rng.choice([90,120,150,180,240]); time = rng.choice([1.5,2,2.5,3,4])
    speed = distance/time
    prompt = f"A vehicle travels {distance} km in {time:g} hours. Find its average speed in km/h."
    return _q(track, topic, diff, prompt, "Use average speed = total distance ÷ total time", ["Write the speed formula.", "Divide the total distance by the total time."], [f"Average speed = {distance} ÷ {time:g} = {speed:g} km/h"], f"{speed:g} km/h", "numeric", speed, tol=1e-6, rng=rng)


def _gen_n5(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    variant = rng.choice(["expand", "factor", "substitute", "nth"])
    x = sp.Symbol('x')
    if variant == "expand":
        a = rng.randint(2,7); b = rng.randint(1,9); c = rng.randint(1,5)
        expr = sp.expand(a*(x+b)-c*x)
        prompt = f"Simplify {a}(x + {b}) - {c}x."
        return _q(track, topic, diff, prompt, "Expand brackets and collect like terms", ["Multiply the factor by every term inside the bracket.", "Then collect the x-terms and constants separately."], [f"{a}(x + {b}) - {c}x = {a}x + {a*b} - {c}x", f"= {sp.sstr(expr)}"], sp.sstr(expr), "expression", sp.sstr(expr), rng=rng)
    if variant == "factor":
        a = rng.randint(2,8); b = rng.randint(2,9); expr = a*x + a*b
        prompt = f"Factorise {sp.sstr(expr)} fully."
        ans = a*(x+b)
        return _q(track, topic, diff, prompt, "Extract a common factor", ["Find the greatest common factor of both terms.", f"Take {a} outside the bracket."], [f"{sp.sstr(expr)} = {a}(x + {b})"], sp.sstr(ans), "expression", sp.sstr(ans), rng=rng)
    if variant == "substitute":
        a,b,c = rng.randint(2,6),rng.randint(1,5),rng.randint(1,8); xv=rng.randint(-3,6)
        val = a*xv*xv+b*xv-c
        prompt = f"Evaluate {a}x^2 + {b}x - {c} when x = {xv}."
        return _q(track, topic, diff, prompt, "Substitute a value into an algebraic expression", ["Replace every x with the given value, using brackets for a negative value.", "Apply powers before multiplication and addition/subtraction."], [f"{a}({xv})^2 + {b}({xv}) - {c}", f"= {val}"], str(val), "numeric", val, rng=rng)
    d = rng.randint(2,8); start = rng.randint(1,10); # arithmetic sequence
    n = sp.Symbol('n'); nth = sp.expand(start+(n-1)*d)
    seq = [start+i*d for i in range(4)]
    prompt = f"Find the nth term of the sequence {', '.join(map(str,seq))}, ..."
    return _q(track, topic, diff, prompt, "Recognise a linear sequence and form its nth term", [f"The common difference is {d}.", "Start with dn and adjust the constant so n = 1 gives the first term."], [f"Common difference = {d}", f"nth term = {sp.sstr(nth)}"], sp.sstr(nth), "expression", sp.sstr(nth), rng=rng)


def _gen_n6(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    m = rng.choice([-4,-3,-2,2,3,4]); c = rng.randint(-6,6); x = rng.randint(-4,6); y=m*x+c
    prompt = f"For the line y = {m}x {'+' if c>=0 else '-'} {abs(c)}, find y when x = {x}."
    return _q(track, topic, diff, prompt, "Evaluate a linear function", ["Substitute the given x-value into the equation.", "Multiply first, then add or subtract the constant."], [f"y = {m}({x}) {'+' if c>=0 else '-'} {abs(c)}", f"y = {y}"], f"y = {y}", "numeric", y, rng=rng)


def _gen_n7(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    if diff == "Stretch" and track in {"O","NA","NT"}:
        r1, r2 = rng.sample(range(-7,8),2)
        b = -(r1+r2); c = r1*r2
        prompt = f"Solve x^2 {'+' if b>=0 else '-'} {abs(b)}x {'+' if c>=0 else '-'} {abs(c)} = 0."
        eqtxt = f"x^2 {'+' if b>=0 else '-'} {abs(b)}x {'+' if c>=0 else '-'} {abs(c)} = 0"
        return _q(track, topic, diff, prompt, "Solve a quadratic equation", ["Bring all terms to one side if necessary.", "Try factorisation or the quadratic formula."], [f"(x - ({r1}))(x - ({r2})) = 0", f"x = {r1} or x = {r2}"], f"x = {r1} or x = {r2}", "quadratic_roots", (r1,r2), metadata={"equations":[eqtxt]}, rng=rng)
    a = rng.randint(2,7); sol = rng.randint(-6,9); b = rng.randint(-10,10); rhs = a*sol+b
    sign = "+" if b>=0 else "-"
    eqtxt = f"{a}x {sign} {abs(b)} = {rhs}"
    prompt = f"Solve {eqtxt}."
    return _q(track, topic, diff, prompt, "Solve a linear equation while preserving equivalence", ["Undo the constant term first.", f"Then divide by the coefficient {a}."], [eqtxt, f"{a}x = {rhs-b}", f"x = {sol}"], f"x = {sol}", "equation_solution", sol, metadata={"equations":[eqtxt]}, rng=rng)


def _gen_n8(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    U = set(range(1,11)); A=set(rng.sample(range(1,11),5)); B=set(rng.sample(range(1,11),5))
    if rng.random()<0.5:
        expected=sorted(A & B); op="A ∩ B"; desc="intersection"
    else:
        expected=sorted(A | B); op="A ∪ B"; desc="union"
    prompt=f"Let A = {{{', '.join(map(str,sorted(A)))}}} and B = {{{', '.join(map(str,sorted(B)))}}}. Find {op}."
    return _q(track, topic, diff, prompt, f"Find the {desc} of two sets", ["Intersection means elements common to both sets; union means elements in either set.", "List each element once."], [f"{op} = {{{', '.join(map(str,expected))}}}"], "{"+", ".join(map(str,expected))+"}", "set", expected, rng=rng)


def _gen_n9(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    A=[[rng.randint(-3,6) for _ in range(2)] for _ in range(2)]; k=rng.randint(2,5); E=[[k*v for v in row] for row in A]
    prompt=f"Given A = [[{A[0][0]}, {A[0][1]}], [{A[1][0]}, {A[1][1]}]], find {k}A."
    return _q(track, topic, diff, prompt, "Multiply a matrix by a scalar", ["Multiply every entry by the scalar.", f"Multiply each of the four entries by {k}."], [f"{k}A = {E}"], str(E), "matrix", E, rng=rng)


def _gen_g1(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    n = rng.choice([5,6,8,10,12]); interior=(n-2)*180; each=interior/n
    prompt=f"Find each interior angle of a regular {n}-sided polygon."
    return _q(track, topic, diff, prompt, "Use the interior-angle sum of a polygon", ["The sum of interior angles is (n − 2) × 180°.", "For a regular polygon, divide the total by n."], [f"Sum = ({n} - 2) × 180° = {interior}°", f"Each angle = {interior} ÷ {n} = {each:g}°"], f"{each:g}°", "numeric", each, tol=0.05, rng=rng)


def _gen_g2(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    small=rng.choice([3,4,5,6]); large=small*rng.choice([1.5,2,2.5]); other=rng.choice([4,5,7,8]); ans=other*large/small
    prompt=f"Two triangles are similar. A side of {small:g} cm in the smaller triangle corresponds to {large:g} cm in the larger triangle. Another side of the smaller triangle is {other:g} cm. Find the corresponding side of the larger triangle."
    return _q(track, topic, diff, prompt, "Use corresponding sides in similar figures", ["Find the linear scale factor from smaller to larger.", "Multiply the second smaller side by the same scale factor."], [f"Scale factor = {large:g} ÷ {small:g} = {large/small:g}", f"Required side = {other:g} × {large/small:g} = {ans:g} cm"], f"{ans:g} cm", "numeric", ans, tol=0.01, rng=rng)


def _gen_g3(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    angle=rng.choice([28,34,42,56]); centre=2*angle
    prompt=f"An angle subtended by the same chord at the circumference of a circle is {angle}°. Find the angle subtended by that chord at the centre."
    return _q(track, topic, diff, prompt, "Apply the circle theorem: angle at centre is twice angle at circumference", ["Recall the relationship between the centre angle and circumference angle standing on the same chord.", "Double the circumference angle."], [f"Centre angle = 2 × {angle}° = {centre}°"], f"{centre}°", "numeric", centre, tol=0.05, rng=rng)


def _gen_g4(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    if diff == "Stretch" and track in {"O","NA"}:
        a=rng.choice([7,8,10,12]); b=rng.choice([9,11,13,15]); C=rng.choice([35,50,65,110]); c=math.sqrt(a*a+b*b-2*a*b*math.cos(math.radians(C)))
        prompt=f"In triangle ABC, sides adjacent to angle C are {a} cm and {b} cm, and C = {C}°. Find the side opposite C."
        return _q(track, topic, diff, prompt, "Use the cosine rule", ["This is not necessarily a right triangle.", "Use c² = a² + b² − 2ab cos C."], [f"c² = {a}² + {b}² − 2({a})({b})cos({C}°)", f"c = {c:.3f} cm"], f"{c:.3f} cm", "numeric", c, tol=0.01, rng=rng)
    opp=rng.choice([3,4,5,6,8]); adj=rng.choice([4,5,7,9,12]); angle=math.degrees(math.atan2(opp,adj))
    prompt=f"In a right-angled triangle, the side opposite angle θ is {opp} cm and the adjacent side is {adj} cm. Find θ."
    return _q(track, topic, diff, prompt, "Use tangent in a right-angled triangle", ["tan θ = opposite ÷ adjacent.", "Use the inverse tangent function on your calculator."], [f"tan θ = {opp}/{adj}", f"θ = tan⁻¹({opp}/{adj}) = {angle:.1f}°"], f"{angle:.1f}°", "numeric", angle, tol=0.15, rng=rng)


def _gen_g5(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    r=rng.choice([3,4,5,6,7]); h=rng.choice([8,10,12,15]); vol=math.pi*r*r*h
    prompt=f"Find the volume of a cylinder of radius {r} cm and height {h} cm. Give your answer to 3 significant figures."
    return _q(track, topic, diff, prompt, "Use V = πr²h and round appropriately", ["Write the cylinder volume formula.", "Substitute the radius and height before rounding."], [f"V = π({r})²({h})", f"V = {vol:.6f} cm³", f"V = {vol:.3g} cm³ (3 s.f.)"], f"{vol:.3g} cm³", "numeric", float(f"{vol:.3g}"), tol=max(0.01, abs(vol)*0.002), rng=rng)


def _gen_g6(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    x1,y1=rng.randint(-5,2),rng.randint(-6,5); x2=x1+rng.choice([2,3,4,5]); m=rng.choice([-3,-2,-1,1,2,3]); y2=y1+m*(x2-x1)
    prompt=f"Find the gradient of the line through ({x1}, {y1}) and ({x2}, {y2})."
    return _q(track, topic, diff, prompt, "Use gradient = change in y ÷ change in x", ["Subtract y-coordinates in the same order as x-coordinates.", "m = (y₂ − y₁)/(x₂ − x₁)."], [f"m = ({y2} - ({y1})) / ({x2} - ({x1}))", f"m = {m}"], f"{m}", "numeric", m, rng=rng)


def _gen_g7(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    a=(rng.randint(-5,6),rng.randint(-5,6)); b=(rng.randint(-5,6),rng.randint(-5,6)); k=rng.randint(2,4); ans=(a[0]+k*b[0],a[1]+k*b[1])
    prompt=f"Given a = ({a[0]}, {a[1]}) and b = ({b[0]}, {b[1]}), find a + {k}b."
    return _q(track, topic, diff, prompt, "Add vectors component-wise and multiply by a scalar", ["Multiply both components of b by the scalar first.", "Then add corresponding x- and y-components."], [f"{k}b = ({k*b[0]}, {k*b[1]})", f"a + {k}b = ({ans[0]}, {ans[1]})"], f"({ans[0]}, {ans[1]})", "pair", ans, rng=rng)


def _gen_s1(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    data=[rng.randint(3,20) for _ in range(rng.choice([6,8,10]))]
    if diff=="Stretch" and track in {"O","NA"}:
        mean=sum(data)/len(data); sd=statistics.pstdev(data)
        prompt=f"For the data {data}, find the population standard deviation. Give your answer to 3 significant figures."
        return _q(track, topic, diff, prompt, "Calculate standard deviation for ungrouped data", ["Find the mean first.", "Use the population standard-deviation formula (divide by n)."], [f"Mean = {mean:.6g}", f"Standard deviation = {sd:.6g}", f"= {sd:.3g} (3 s.f.)"], f"{sd:.3g}", "numeric", float(f"{sd:.3g}"), tol=max(0.005, sd*0.002), rng=rng)
    med=statistics.median(data)
    prompt=f"Find the median of the data: {', '.join(map(str,data))}."
    return _q(track, topic, diff, prompt, "Order data and identify the middle value(s)", ["Arrange the values from smallest to largest.", "If there are two middle values, average them."], [f"Ordered data: {sorted(data)}", f"Median = {med:g}"], f"{med:g}", "numeric", med, rng=rng)


def _gen_s2(track: str, topic: Topic, diff: str, rng: random.Random) -> Question:
    red=rng.randint(2,8); blue=rng.randint(2,8); green=rng.randint(1,6); total=red+blue+green
    if diff=="Stretch" and track in {"O","NA"}:
        # two independent draws with replacement
        p=(red/total)*(red/total)
        prompt=f"A bag has {red} red, {blue} blue and {green} green counters. A counter is drawn, replaced, then another is drawn. Find the probability of drawing red twice."
        return _q(track, topic, diff, prompt, "Multiply probabilities for independent events", ["Replacement means the probabilities stay the same for the second draw.", "Multiply P(red) by P(red)."], [f"P(red) = {red}/{total}", f"P(red twice) = ({red}/{total})² = {Fraction(red*red,total*total)}"], str(Fraction(red*red,total*total)), "numeric", p, tol=1e-6, rng=rng)
    p=red/total
    prompt=f"A bag contains {red} red, {blue} blue and {green} green counters. One counter is chosen at random. Find P(red)."
    frac=Fraction(red,total)
    return _q(track, topic, diff, prompt, "Probability = favourable outcomes ÷ total outcomes", ["Count the total number of counters.", "Use number of red counters divided by total counters."], [f"Total = {total}", f"P(red) = {red}/{total} = {frac}"], str(frac), "numeric", p, tol=1e-6, rng=rng)


GENERATORS: dict[str, Callable[[str, Topic, str, random.Random], Question]] = {
    "N1": _gen_n1,
    "N2": _gen_n2,
    "N3": _gen_n3,
    "N4": _gen_n4,
    "N5": _gen_n5,
    "N6": _gen_n6,
    "N7": _gen_n7,
    "N8": _gen_n8,
    "N9": _gen_n9,
    "G1": _gen_g1,
    "G2": _gen_g2,
    "G3": _gen_g3,
    "G4": _gen_g4,
    "G5": _gen_g5,
    "G6": _gen_g6,
    "G7": _gen_g7,
    "S1": _gen_s1,
    "S2": _gen_s2,
}


def analyze_own_algebra_question(question_text: str, working_text: str) -> AttemptResult:
    """Best-effort free-form checker for a single-variable algebra equation.

    It intentionally refuses unsupported natural-language/diagram questions rather
    than pretending to understand them offline.
    """
    try:
        eq = _parse_equation(question_text)
    except Exception as exc:
        raise ValueError(
            "Offline free-form checking currently supports typed algebra questions containing an equation, for example: Solve 3(x + 2) = 18. For other topics, use Generated Practice so the tutor knows the exact mathematical structure."
        ) from exc
    syms = list(eq.free_symbols)
    if len(syms) != 1:
        raise ValueError("Offline free-form checking currently supports one-variable equations only.")
    var = syms[0]
    sols = sp.solve(eq, var)
    if not sols:
        raise ValueError("I could not obtain a finite solution for this equation offline.")
    if len(sols) == 1:
        checker="equation_solution"; expected=float(sp.N(sols[0])); answer=f"{var} = {sp.sstr(sols[0])}"
    else:
        checker="quadratic_roots"; expected=tuple(float(sp.N(s)) for s in sols); answer=" or ".join(f"{var} = {sp.sstr(s)}" for s in sols)
    topic = next(t for t in TOPICS if t.code == "N7")
    q = Question(
        id="own-algebra",
        track="CUSTOM",
        topic_code="N7",
        topic_name=topic.name,
        strand=topic.strand,
        difficulty="Student question",
        prompt=question_text,
        target_skill="Preserve equivalence while solving the equation",
        hints=["Identify the inverse operation that simplifies the equation without changing its solution set.", "Check the first line that is no longer equivalent to the original equation.", f"Verified solution: {answer}"],
        worked_solution=[f"Verified solution: {answer}"],
        answer_display=answer,
        checker=checker,
        expected=expected,
        metadata={"equations":[str(eq.lhs)+" = "+str(eq.rhs)]},
    )
    return evaluate_attempt(q, working_text)
