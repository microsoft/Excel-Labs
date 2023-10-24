/*

Spreadsheet Lisp (SL) v0.7.0

Forged in the Advanced Formula Environment (AFE),
courtesy of the minds on the Excel Labs team of Microsoft Garage.

SL bows respectfully to its ancestral list of LISPs:
(1.5 Common Scheme Racket Pico et.al.)

DISCLAIMER:
Most SL functions do not require macros to be enabled. [*.xlsx]
The coolest SL functions require macros to be enabled. [*.xlsm]
You are thus free, within the bounds of determinism, to choose your path forward.

"For a dreamer is one who can only find his way by moonlight,
 and his punishment is that he sees the dawn before the rest of the world."
 - Oscar Wilde

*/

TABLE_OF_CONTENTS =
    VLIST(
        "Table of Contents",
        "Types",
        "Functions [.xlsm]",
        "Functions [.xlsx]",
        "Arcade",
        "Epilogue");

/*
Table of Contents

Types:
Booleans
Atoms

Macro-dependent functions [*.xlsm]:
=EVAL(expression)
=APPLY(functor, context)
=CURRY(reference, free_variable)

Macro-free functions [*.xlsx]:
=L(functor_string, [argument], ...)
=ARITY([argument], ...)
=IS(argument)
=NOTERROR(value)
=PROVIDED(argument)
=BOTHPROVIDED(argument1, argument2)
=ISTRUE(value)
=ISFALSE(value)
=MAYBE(input)
=MAYBENOT(potential, otherwise)
=UNKNOWN(input)
=KNOWN(input)
=MISSING(input)
=SOMETHING(input)
=TRUTHY(value)
=FALSY(value)
=ONE(value)
=ID(x)
=CAR(range)
=_FIRST(range)
=_SECOND(range)
=_THIRD(range)
=_FOURTH(range)
=_FIFTH(range)
=CDR(range)
=REST(range)
=APPEND(head, body)
=CONS(head, body)
=EQ(range)
=EQUAL(a, b)
=DIV(numerator, denominator)
=DIFFERENCE(minuend, subtrahend)
=VLIST([a], ...)
=HLIST([a], ...)
=MEMBER(needle, haystack)
=ATOM(value)
=ATOMIC(value)
=FIRSTROW(range)
=FIRSTCOL(range)
=NEGATE(range)
=DECREMENT(number)
=INCREMENT(number)
=CONTAINS(haystack, needle)
=NOTBLANK(value)
=HOMOGENOUS(range)
=HETEROGENOUS(range)
=SHEETNAME()
=GET(range, [sheet], [file], [path])
=FORMAT(template, [first], ...)
=SQL(table, query)
=SUBSET(table, query)
=GETCOL(index, range)
=GETCOLS(indexes, range)
=FIRSTWORD(sentence)
=FIRSTLETTER(text)
=COLINDEXES(subset, superset)
=TRIMALL(range)
=GREATERTHAN(x, y)
=GTE(x, y)
=LESSTHAN(x, y)
=LTE(x, y)
=ORIENTATION(range)
=CAGR(beginning_value, ending_value, [periods])
=FACTORIAL(n)
=ALPHABET([vertical])
=LETTER(reference)
=NEITHER(this, that)
=POSITIVE(value)
=NEGATIVE(value)
=NONZERO(value)
=ZERO(value)
=HORIZONTAL(range)
=VERTICAL(range)
=PLURAL(value)
=TYPESTRING(value, [recur])
=ALL(truth_values)
=CASE(success, attempt)
=TYPEDISPATCH(context, case1, result1, [case2], [result2], ...)
=LENGTH(input)
=NOTFALSE(condition)
=NOTEQUAL(this, that)
=BEFORE(index, range)
=AFTER(index, range)
=MAYBEAPPEND(head, body)
=WITHOUT(range, index_or_string)
=UNITLIST(list)
=COLNUM(column_string)
=LISTINDEX(list, index)

Arcade:
=FLIPCOIN([times])
=MAGIC8BALL([question])
=ROCKPAPERSCISSORS(throw)

Epilogue:
Fledgling Style Guide

*/

///
/// Types
///

// Booleans
else = TRUE;
t = TRUE;
f = FALSE;
null = #NULL!;

// Atoms
ATOMS =
    HLIST("number", "string", "boolean", "error");

ISATOM =
    LAMBDA(
        thing,
        MEMBER(TYPESTRING(thing), ATOMS));

ISRATIONAL =
    LAMBDA(
        value,
        NOTERROR(DIV(value, 1)));

///
/// Macro-dependent Functions
///

// [A1]: '(sum 1 2 3 4 5)
// =EVAL(A1) -> 15
EVAL =
    LAMBDA(
        expression,
        LET(
            expr,
                TRIM(expression),
            fulcrum,
                FIND(" ", expr),
            characters,
                LEN(expr),
            functor,
                MID(expr, 2, fulcrum-2),
            context,
                RIGHT(expr, characters-fulcrum),
            csv,
                SUBSTITUTE(context, " ", ", "),
            formula,
                FORMAT("={1}({2}", functor, csv),
            VALUE(EVALUATE(formula))));

// APPLY currently only accepts two strings
// =APPLY("sum", "1,2,3") -> 6
APPLY =
    LAMBDA(
        functor, context,
        EVAL(
            FORMAT(
                "({1} {2})",
                functor,
                TRIM(SUBSTITUTE(context, ",", " ")))));

// [A1] =SUM(1, 2) -> 3
// =CURRY(A1, 3) -> 6
CURRY =
    LAMBDA(
        reference, free_variable,
        LET(
            formula, FORMULATEXT(reference),
            without_closing_paren,
                LEFT(formula, DECREMENT(LEN(formula))),
            EVALUATE(
                FORMAT(
                    "{1}, {2})",
                    without_closing_paren,
                    free_variable))));
///
/// Functions
///

// L/1[4]: The Lisp Function is a middle-way approach between Excel Formulas and Lisp Forms.
// =L("sum", 1, 2, 3) -> 6
// =L("product", 5, 5) -> 25 
L =
    LAMBDA(
        functor_string, [_1], [_2], [_3], [_4],
        LET(
            arity, ARITY(functor_string, _1, _3, _3, _4),
            _functor, TEXT(functor_string, "0"),
            SWITCH(
                arity,
                1, EVAL(FORMAT("({1})", _functor)),
                2, EVAL(FORMAT("({1} {2})", _functor, _1)),
                3, EVAL(FORMAT("({1} {2} {3})", _functor, _1, _2)),
                4, EVAL(FORMAT("({1} {2} {3} {4})", _functor, _1, _2, _3)),
                5, EVAL(FORMAT("({1} {2} {3} {4} {5})", _functor, _1, _2, _3, _4)))));

// ARITY/[9]: The Arity function determines how many values were provided to a given function.
// =ARITY(0/0) -> 1
// =ARITY(A1:A5) -> 1
// =ARITY(1, 2, 3) -> 3
// =ARITY("a", "b", "c", "d") -> 4
ARITY =
    LAMBDA(
        [_1], [_2], [_3], [_4], [_5], [_6], [_7], [_8], [_9],
        SUM(IS(_1), IS(_2), IS(_3), IS(_4), IS(_5), IS(_6), IS(_7), IS(_8), IS(_9)));

// IS/1: The Is helper function aids ARITY() by determining whether an optional argument was provided.
IS =
    LAMBDA(
        argument,
        IF(
            ISOMITTED(argument),
            0,
            1));

// =NOTERROR(123) -> TRUE
// =NOTERROR(#N/A!) -> FALSE
NOTERROR =
    LAMBDA(
        value,
        NOT(ISERROR(value)));

PROVIDED =
    LAMBDA(
        input,
        AND(
            NOT(ISOMITTED(input))));

BOTHPROVIDED =
    LAMBDA(
        input1, input2,
        AND(
            NOT(ISOMITTED(input1)),
            NOT(ISOMITTED(input2))));

// =ISTRUE(TRUE) -> TRUE
// =ISTRUE(FALSE) -> FALSE
ISTRUE =
    LAMBDA(
        value,
        AND(
            TYPESTRING(value)="boolean",
            EQUAL(value, t)));

// =ISFALSE(TRUE) -> FALSE
// =ISFALSE(FALSE) -> TRUE
ISFALSE =
    LAMBDA(
        value,
        AND(
            TYPESTRING(value)="boolean",
            EQUAL(value, f)));

// =MAYBE(123) -> 123
// =MAYBE(TRUE) -> TRUE
// =MAYBE(#NULL!) -> FALSE
// =MAYBE([missing_argument]) -> FALSE
MAYBE =
    LAMBDA(
        input,
        IF(
            AND(
                NOTERROR(input),
                PROVIDED(input)),
            input,
            ERROR.TYPE(12)));

// =MAYBENOT(0, 1) -> 0
// =MAYBENOT(0/0, 1) -> 1
// =MAYBENOT("Provided", "Missing") -> "Provided"
// =MAYBENOT(, "Missing") -> "Missing"
MAYBENOT =
    LAMBDA(
        potential, otherwise,
        IF(
            SOMETHING(potential),
            potential,
            otherwise));

UNKNOWN =
    LAMBDA(
        input,
        IF(
            ISERROR(input),
            EQUAL(ERROR.TYPE(input), 12)));

KNOWN =
    LAMBDA(
        input,
        NOT(UNKNOWN(input)));

MISSING =
    LAMBDA(
        input,
        ISOMITTED(input));

SOMETHING =
    LAMBDA(
        input,
        AND(
            NOTERROR(input),
            PROVIDED(input)));

// =TRUTHY(1) -> TRUE
// =TRUTHY(2) -> FALSE
// =TRUTHY(TRUE) -> TRUE
TRUTHY =
    LAMBDA(
        value,
        OR(
            EQUAL(value, 1),
            EQUAL(value, t),
            EQUAL(UPPER(value), "TRUE")));

FALSY =
    LAMBDA(
        value,
        MEMBER(
            MAYBE(value),
            HLIST("", f, 0)));

// =ONE(0) -> FALSE
// =ONE(1) -> TRUE
// =ONE(2) -> FALSE
ONE =
    LAMBDA(
        value,
        EQUAL(1, value));

// =ID(123) -> 123
ID =
    LAMBDA(
        x,
        x);

// =CAR(A1:A5) -> A1
CAR =
    LAMBDA(
        range,
        INDEX(range, 1, 1));

// =_FIRST(A1:A5) -> A1
_FIRST =
    LAMBDA(
        range,
        CAR(range));

// =_SECOND(A1:A5) -> A2
_SECOND =
    LAMBDA(
        range,
        CAR(CDR(range)));

// =_THIRD(A1:A5) -> A3
_THIRD =
    LAMBDA(
        range,
        CAR(CDR(CDR(range))));

// =_FOURTH(A1:A5) -> A4
_FOURTH =
    LAMBDA(
        range,
        CAR(CDR(CDR(CDR(range)))));

// =_FIFTH(A1:A5) -> A5
_FIFTH =
    LAMBDA(
        range,
        CAR(CDR(CDR(CDR(CDR(range))))));

// =CDR(A1:A5) -> A2:A5
CDR =
    LAMBDA(
        range,
        LET(
            horizontal, COLUMNS(range)>ROWS(range),
            IF(
                horizontal,
                MAKEARRAY(
                    1,
                    DECREMENT(COLUMNS(range)),
                    LAMBDA(
                        _row, col,
                        INDEX(range, 1, INCREMENT(col)))),
                MAKEARRAY(
                    DECREMENT(ROWS(range)),
                    1,
                    LAMBDA(
                        row, _col,
                        INDEX(range, INCREMENT(row), 1))))));

REST =
    LAMBDA(
        range,
        CDR(range));

// =APPEND(A1:A5, A6:A10) -> A1:A10
APPEND =
    LAMBDA(
        head, body,
        LET(
            horizontal, COLUMNS(body)>ROWS(body),
            IF(
                horizontal,
                HSTACK(head, body),
                VSTACK(head, body))));

// =CONS(1, {2,3}) -> {1,2,3}
CONS =
    LAMBDA(
        head, body,
        IF(
            ATOMIC(head),
            APPEND(head, body)));

// =EQ({1,1,1}) -> TRUE
// =EQ({1,1,2}) -> FALSE
EQ =
    LAMBDA(
        range,
        AND(EXACT(CAR(range), range)));

// =EQUAL(1, 1) -> TRUE
// =EQUAL(1, 2) -> FALSE
EQUAL =
    LAMBDA(
        a, b,
        a=b);

DIV =
    LAMBDA(
        numerator, denominator,
        numerator / denominator);

DIFFERENCE =
    LAMBDA(
        minuend, subtrahend,
        minuend-subtrahend);

// Builds a single-column (vertical) range
// Syntactic sugar for {1;2;3;4;5;...}
VLIST =
    LAMBDA(
        [_1], [_2], [_3], [_4], [_5], [_6], [_7], [_8], [_9], [_10],
        [_11], [_12], [_13], [_14], [_15], [_16], [_17], [_18], [_19], [_20],
        LET(
            row_count,
                SUM(
                    IS(_1), IS(_2), IS(_3), IS(_4), IS(_5), IS(_6), IS(_7),
                    IS(_8), IS(_9), IS(_10), IS(_11), IS(_12), IS(_13), IS(_14),
                    IS(_15), IS(_16), IS(_17), IS(_18), IS(_19), IS(_20)),
            MAKEARRAY(
                row_count,
                1,
                LAMBDA(
                    row,
                    _col,
                    CHOOSE(
                        row,
                        _1, _2, _3, _4, _5, _6, _7, _8, _9, _10,
                        _11, _12, _13, _14, _15, _16, _17, _18, _19, _20)))));

// Builds a single-row (horizontal) range
// Syntactic sugar for {1,2,3,4,5,...}
HLIST =
    LAMBDA(
        [_1], [_2], [_3], [_4], [_5], [_6], [_7], [_8], [_9], [_10],
        [_11], [_12], [_13], [_14], [_15], [_16], [_17], [_18], [_19], [_20],
        LET(
            column_count,
                SUM(
                    IS(_1), IS(_2), IS(_3), IS(_4), IS(_5), IS(_6), IS(_7),
                    IS(_8), IS(_9), IS(_10), IS(_11), IS(_12), IS(_13), IS(_14),
                    IS(_15), IS(_16), IS(_17), IS(_18), IS(_19), IS(_20)),
            MAKEARRAY(
                1,
                column_count,
                LAMBDA(
                    _row,
                    col,
                    CHOOSE(
                        col,
                        _1, _2, _3, _4, _5, _6, _7, _8, _9, _10,
                        _11, _12, _13, _14, _15, _16, _17, _18, _19, _20)))));

// =MEMBER(1, {1,2,3}) -> TRUE
// =MEMBER(1, {2,3,4}) -> FALSE
MEMBER =
    LAMBDA(
        needle, haystack,
        OR(EXACT(needle, haystack)));

// =ATOM(123) -> TRUE
// =ATOM("ABC") -> TRUE
// =ATOM(FALSE) -> TRUE
// =ATOM({1,2,3}) -> FALSE
// =ATOM(#REF!) -> FALSE
ATOM =
    LAMBDA(
        value,
        LET(
            number, 1,
            text, 2,
            boolean, 4,
            atoms, HLIST(number, text, boolean),
            MEMBER(TYPE(value), atoms)));

// =ATOMIC(#REF!) -> TRUE
ATOMIC =
    LAMBDA(
        value,
        OR(
            ISERROR(value),
            ATOM(value)));

// =FIRSTROW(A1:D10) -> A1:D1
FIRSTROW =
    LAMBDA(
        range,
        INDEX(range, 1, 0));

// =FIRSTCOL(A1:D10) -> A1:A10
FIRSTCOL =
    LAMBDA(
        range,
        INDEX(range, 0, 1));

// =NEGATE(123) -> -123
NEGATE =
    LAMBDA(
        range,
        MAP(range, LAMBDA(x, -x)));

// =DECREMENT(99) -> 98
DECREMENT =
    LAMBDA(
        number,
        SUM(number, NEGATE(1)));

// INCREMENT(0) -> 1
INCREMENT =
    LAMBDA(
        number,
        SUM(number, 1));

// =CONTAINS({1,2,3}, 3) -> TRUE
// =CONTAINS({4,5,6}, 3) -> FALSE
CONTAINS =
    LAMBDA(
        haystack,
        needle,
        LET(
            two_strings, AND(ISTEXT(haystack), ISTEXT(needle)),
            IF(
                two_strings,
                NOT(ISERROR(FIND(needle, haystack))),
                OR(EXACT(needle, haystack)))));

NOTBLANK =
    LAMBDA(
        value,
        NOTEQUAL(value, ""));

HOMOGENOUS =
    LAMBDA(
        range,
        EQ(range));

HETEROGENOUS =
    LAMBDA(
        range,
        NOT(HOMOGENOUS(range)));

// Workbook must be saved to access "filename"
SHEETNAME =
    LAMBDA(
        LET(
            full_name, CELL("filename", INDIRECT(CONCAT("A1"))),
            full_length, LEN(full_name),
            path_length, FIND("]", full_name),
            reference_length, full_length - path_length,
            reference, RIGHT(full_name, reference_length),
            exclamation_index, IFERROR(FIND("!", reference), 0),
            IF(
                exclamation_index > 0,
                LEFT(reference, exclamation_index),
                reference)));

// =GET("B2", "Sheet3") -> =Sheet3!B2
GET =
    LAMBDA(
        range, [sheet], [file], [path],
        LET(
            sheet_provided, PROVIDED(sheet),
            file_provided, PROVIDED(file),
            path_provided, PROVIDED(path),
            reference_wrapper, "'",
            INDIRECT(
                CONCAT(
                    IF(
                        file_provided,
                        CONCAT(
                            reference_wrapper,
                            IF(path_provided, path, ""),
                            "[", file, "]"),
                        ""),
                    IF(
                        sheet_provided,
                        CONCAT(
                            IF(file_provided, "", reference_wrapper),
                            sheet, reference_wrapper, "!"),
                        ""),
                    range))));

// =FORMAT("The time is now {1}:{2}:{2}.", 11, 59) -> "The time is now 11:59:59."
FORMAT =
    LAMBDA(
        template, [first], [second], [third], [fourth], [fifth],
        LET(
            _after1, IF(PROVIDED(first), SUBSTITUTE(template, "{1}", first), template),
            _after2, IF(PROVIDED(second), SUBSTITUTE(_after1, "{2}", second), _after1),
            _after3, IF(PROVIDED(third), SUBSTITUTE(_after2, "{3}", third), _after2),
            _after4, IF(PROVIDED(fourth), SUBSTITUTE(_after3, "{4}", fourth), _after3),
            IF(PROVIDED(fifth), SUBSTITUTE(_after4, "{5}", fifth), _after4)));

// =SQL(range, "select [col], [etc.] where [col] >= 23")
// Currently only supports single-condition where clause.
SQL =
    LAMBDA(
        table, query,
        LET(
            first_word, FIRSTWORD(query),
            rest_of_query, TRIM(TEXTAFTER(query, first_word)),
            SWITCH(
                first_word,
                "select", SUBSET(table, rest_of_query),
                else, table)));

SUBSET =
    LAMBDA(
        table, query,
        LET(
            where_index, IFERROR(FIND("where", LOWER(query)), 0),
            source_columns, FIRSTROW(table),
            column_string, TRIM(IF(where_index>0, LEFT(query, where_index-1), query)),
            queried_columns, TRIMALL(TEXTSPLIT(column_string, ",")),
            column_indexes, COLINDEXES(queried_columns, source_columns),
            subset, GETCOLS(table, column_indexes),
            IF(
                where_index = 0,
                subset,
                LET(
                    where_clause, TEXTSPLIT(TRIM(TEXTAFTER(query, "where")), " "),
                    column_header, _FIRST(where_clause),
                    condition, _SECOND(where_clause),
                    value, VALUE(_THIRD(where_clause)),
                    column_index, COLINDEXES(column_header, queried_columns),
                    filtered_column, GETCOLS(subset, column_index),
                    VSTACK(
                        queried_columns,
                        FILTER(
                            subset,
                            SWITCH(
                                condition,
                                "=", filtered_column=value,
                                ">", filtered_column>value,
                                ">=", filtered_column>=value,
                                "<", filtered_column<value,
                                "<=", filtered_column<=value)))))));

// =GETCOL(2, A1:C10) -> B1:B10
GETCOL =
    LAMBDA(
        index, range,
        INDEX(range, 0, index));

// =GETCOLS(A1:D10, {2,3}) -> B1:C10
GETCOLS =
    LAMBDA(
        indexes, range,
        IF(
            COLUMNS(indexes) = 1,
            GETCOL(CAR(indexes), range),
            HSTACK(
                GETCOL(CAR(indexes), range),
                GETCOLS(range, CDR(indexes)))));

// =FIRSTWORD("The truth is rarely pure and never simple.") -> "The"
FIRSTWORD =
    LAMBDA(
        sentence,
        LEFT(sentence, TRIM(FIND(" ", sentence)-1)));

// =FIRSTLETTER("Spreadsheet") -> "S"
FIRSTLETTER =
    LAMBDA(
        text,
        LEFT(text, 1));

// =COLINDEXES({"b","d"}, {"a","b","c""d","e"}) -> {2,4}
COLINDEXES =
    LAMBDA(
        subset, superset,
        MAKEARRAY(
            1,
            COLUMNS(subset),
            LAMBDA(
                _row, col,
                MATCH(
                    INDEX(subset, 1, col),
                    superset,
                    f))));

// =TRIMALL({"  a ","b  "," c "}) -> {"a","b","c"}
TRIMALL =
    LAMBDA(
        range,
        MAKEARRAY(
            ROWS(range),
            COLUMNS(range),
            LAMBDA(r, c, TRIM(INDEX(range, r, c)))));

// =GREATERTHAN(2, 1) -> TRUE
GREATERTHAN =
    LAMBDA(
        x, y,
        x>y);

// =GTE(2, 2) -> TRUE
GTE =
    LAMBDA(
        x, y,
        x>=y);

// =LESSTHAN(1, 2) -> TRUE
LESSTHAN =
    LAMBDA(
        x, y,
        x<y);

// =LTE(1, 1) -> TRUE
LTE =
    LAMBDA(
        x, y,
        x<=y);

// =ORIENTATION({1,2,3}) -> "Horizontal"
// =ORIENTATION({1;2;3}) -> "Vertical"
// =ORIENTATION(A1:E1) -> "Horizontal"
// =ORIENTATION(A1:A5) -> "Vertical"
ORIENTATION =
    LAMBDA(
        range,
        LET(
            horizontal, GREATERTHAN(COLUMNS(range), ROWS(range)),
            IF(
                horizontal,
                "Horizontal",
                "Vertical")));

// =CAGR(100, 150) -> 0.5
// =CAGR(150, 100) -> -0.333
// =CAGR(100, 200) -> 1
// =CAGR(100, 0) -> -1
CAGR =
    LAMBDA(
        beginning_value, ending_value, [periods],
        LET(
            p, IF(PROVIDED(periods), periods, 1),
            nominal_growth, ending_value / beginning_value,
            DECREMENT(POWER(nominal_growth, 1/p))));

// =FACTORIAL(0) -> 1
// =FACTORIAL(1) -> 1
// =FACTORIAL(2) -> 2
// =FACTORIAL(3) -> 6
// =FACTORIAL(4) -> 24
// =FACTORIAL(5) -> 123
FACTORIAL =
    LAMBDA(
        n,
        IF(
            LTE(n, 1),
            1,
            PRODUCT(n, FACTORIAL(DECREMENT(n)))));

ALPHABET =
    LAMBDA(
        [vertical],
        LET(
            alphabet,
                HSTACK(
                    {"A","B","C","D","E","F"},
                    {"G","H","I","J","K","L"},
                    {"M","N","O","P","Q","R"},
                    {"S","T","U","V","W","X","Y","Z"}),
            IF(
                TRUTHY(MAYBE(vertical)),
                TRANSPOSE(alphabet),
                alphabet)));

// =LETTER(A1) -> "A"
// =LETTER(BC23) -> "BC"
LETTER =
    LAMBDA(
        reference,
        LET(
            letters,  ALPHABET(),
            colnum, COLUMN(reference),
            iterations, FLOOR.MATH(colnum/26, 1),
            remainder, MOD(colnum, 26),
            IF(
                colnum<=26,
                INDEX(letters, 1, colnum),
                CONCAT(
                    INDEX(letters, 1, iterations),
                    INDEX(letters, 1, remainder)))));

// =NEITHER(TRUE, TRUE) -> FALSE
// =NEITHER(TRUE, FALSE) -> FALSE
// =NEITHER(FALSE, TRUE) -> FALSE
// =NEITHER(FALSE, FALSE) -> TRUE
NEITHER =
    LAMBDA(
        this, that,
        AND(NOT(this), NOT(that)));

// POSITIVE(1) -> TRUE
// POSITIVE(0) -> FALSE
// POSITIVE(-1) -> FALSE
POSITIVE =
    LAMBDA(
        value,
        AND(
            ISNUMBER(value),
            GREATERTHAN(value, 0)));

// NEGATIVE(1) -> FALSE
// NEGATIVE(0) -> TRUE
// NEGATIVE(-1) -> TRUE
NEGATIVE =
    LAMBDA(
        value,
        AND(
            ISNUMBER(value),
            GREATERTHAN(0, value)));

// NONZERO(1) -> TRUE
// NONZERO(0) -> FALSE
// NONZERO(-1) -> TRUE
NONZERO =
    LAMBDA(
        value,
        OR(
            POSITIVE(value),
            NEGATIVE(value)));

// ZERO(1) -> TRUE
// ZERO(0) -> TRUE
// ZERO(-1) -> FALSE
ZERO =
    LAMBDA(
        value,
        EQUAL(value, 0));

// =HORIZONTAL(A1:D1) -> TRUE
// =HORIZONTAL(A1:A10) -> FALSE
HORIZONTAL =
    LAMBDA(
        range,
        GREATERTHAN(COLUMNS(range), ROWS(range)));

// =VERTICAL(A1:D1) -> FALSE
// =VERTICAL(A1:A10) -> TRUE
VERTICAL =
    LAMBDA(
        range,
        GREATERTHAN(COLUMNS(range), ROWS(range)));

// =PLURAL(0) -> FALSE
// =PLURAL(1) -> FALSE
// =PLURAL(2) -> TRUE
PLURAL =
    LAMBDA(
        value,
        AND(
            ISNUMBER(value),
            GREATERTHAN(value, 1)));

// =TYPESTRING(1) -> "number"
// =TYPESTRING(TRUE) -> "boolean"
// =TYPESTRING("three") -> "string"
// =TYPESTRING(#REF!) -> "error"
// =TYPESTRING({1,2,3}) -> "dimensional"
TYPESTRING =
    LAMBDA(
        value, [recur],
        LET(
            recursive, TRUTHY(MAYBE(recur)),
            type_number, IF(MEMBER(value, {""}), 0, TYPE(value)),
            SWITCH(
                type_number,
                0, "empty",
                1, "number",
                2, "string",
                4, "boolean",
                16, "error",
                64,
                    IF(
                        recursive,
                        MAP(value, LAMBDA(item, TYPESTRING(item))),
                        "dimensional"),
                128, "compound",
                else, "unknown")));

// =ALL({TRUE,TRUE,TRUE}) -> TRUE
// =ALL({TRUE,TRUE,FALSE}) -> FALSE
ALL =
    LAMBDA(
        truth_values,
        AND(EXACT(t, truth_values)));

CASE =
    LAMBDA(
        success, attempt,
        LET(
            right, attempt,
            left, IF(ISTRUE(attempt), right, success),
            AND(
                EQUAL(ROWS(left), ROWS(right)),
                EQUAL(COLUMNS(left), COLUMNS(right)),
                ALL(EXACT(left, right)))));
/*
=TYPEDISPATCH(
    A1:A3,
    {"number","number","number"}, "numbers",
    {"string","string","string"}, "letters",
    {"string","string","number"}, "mixed",
    else, "other")
*/
TYPEDISPATCH =
    LAMBDA(
        context,
        case1, result1,
        [case2], [result2],
        [case3], [result3],
        [case4], [result4],
        [case5], [result5],
        [case6], [result6],
        [case7], [result7],
        [case8], [result8],
        [case9], [result9],
        [case10], [result10],
        LET(
            types, TYPESTRING(context, t),
            IFS(
                CASE(types, case1), result1,
                CASE(types, MAYBE(case2)), result2,
                CASE(types, MAYBE(case3)), result3,
                CASE(types, MAYBE(case4)), result4,
                CASE(types, MAYBE(case5)), result5,
                CASE(types, MAYBE(case6)), result6,
                CASE(types, MAYBE(case7)), result7,
                CASE(types, MAYBE(case8)), result8,
                CASE(types, MAYBE(case9)), result9,
                CASE(types, MAYBE(case10)), result10)));

// =LENGTH("Lisp") -> 4
// =LENGTH({1,2,3,4,5}) -> 5
LENGTH =
    LAMBDA(
        input,
        TYPEDISPATCH(
            input,
            "string", LEN(input),
            else, COUNTA(input)));

NOTFALSE =
    LAMBDA(
        condition,
        AND(
            EQUAL(TYPESTRING(condition), "boolean"),
            NOT(EQUAL(condition, f))));

NOTEQUAL =
    LAMBDA(
        this, that,
        NOT(EQUAL(this, that)));

// =BEFORE(4, {1,2,3,4,5}) -> {1,2,3}
BEFORE =
    LAMBDA(
        index, range,
        MAKEARRAY(
            IF(HORIZONTAL(range), 1, MIN(DECREMENT(index), LENGTH(range))),
            IF(HORIZONTAL(range), MIN(DECREMENT(index), LENGTH(range)), 1),
            LAMBDA(
                row, col,
                INDEX(range, row, col))));

// =AFTER(2, {1,2,3,4,5}) -> {3,4,5}
AFTER =
    LAMBDA(
        index, list,
        LET(
            idx, MAX(index, 0), 
            MAKEARRAY(
                IF(HORIZONTAL(list), 1, DIFFERENCE(ROWS(list), idx)),
                IF(HORIZONTAL(list), DIFFERENCE(COLUMNS(list), idx), 1),
                LAMBDA(
                    row, col,
                    INDEX(
                        list,
                        IF(HORIZONTAL(list), 1, SUM(row, idx)),
                        IF(HORIZONTAL(list), SUM(col, idx), 1))))));

MAYBEAPPEND =
    LAMBDA(
        head, body,
        LET(
            _head, MAYBE(head),
            _body, MAYBE(body),
            IFS(
                AND(SOMETHING(_head), SOMETHING(_body)),
                    APPEND(_head, _body),
                ISERROR(body), _head,
                else, _body)));

// REMOVE/2: Removes an item at a given index from a given list 
// =REMOVE(3, {5,6,7,8,9}) -> {5,6,8,9}
WITHOUT =
    LAMBDA(
        range, index_or_string,
        LET(
            index,
                IF(
                    ISNUMBER(index_or_string),
                    index_or_string,
                    MATCH(index_or_string, range, 0)),
            MAYBEAPPEND(
                BEFORE(index, range),
                AFTER(index, range))));

// UNITLIST/1: Determines whether a list contains a single value
UNITLIST =
    LAMBDA(
        list,
        ONE(LENGTH(list)));

// =COLNUM(A1) -> 1
// =COLNUM(B1) -> 2
// =COLNUM(Z1) -> 26
// =COLNUM(AA1) -> 27
COLNUM =
    LAMBDA(
        column_string,
        SWITCH(
            LENGTH(column_string),
            1, MATCH(UPPER(column_string), ALPHABET(), 0),
            2, 
            LET(
                first, MATCH(UPPER(LEFT(column_string, 1)), ALPHABET(), 0),
                second, MATCH(UPPER(RIGHT(column_string, 1)), ALPHABET(), 0),
                SUM(26, PRODUCT(26, DECREMENT(first)), second))));

// =LISTINDEX(HLIST("a", "b", "c", "d"), 3) -> "c"
LISTINDEX =
    LAMBDA(
        list, index,
        IF(
            HORIZONTAL(list),
            INDEX(list, 1, index),
            INDEX(list, index, 1)));

///
/// Arcade
///

// =FLIPCOIN() -> "Heads"
// =FLIPCOIN(3) -> {"Heads";"Tails";"Tails"}
FLIPCOIN =
    LAMBDA(
        [times],
        LET(
            _TOSS, LAMBDA(CHOOSE(RANDBETWEEN(1, 2), "Heads", "Tails")),
            _times, IF(PROVIDED(times), times, 1),
            IF(
                LTE(_times, 1),
                _TOSS(),
                VSTACK(
                    _TOSS(),
                    FLIPCOIN(DECREMENT(_times))))));

// =MAGIC8BALL("Is Spreadsheet Lisp fun to use?") -> "It is decidedly so."
MAGIC8BALL =
    LAMBDA(
        [question],
        IF(
            OR(PROVIDED(question), ISBLANK(MAYBE(question))),
            CHOOSE(
                RANDBETWEEN(1, 20),
                "It is certain",
                "Reply hazy, try again",
                "It is decidedly so",
                "Without a doubt",
                "Don't count on it",
                "Yes definitely",
                "Ask again later",
                "You may rely on it",
                "My reply is no",
                "As I see it, yes",
                "Better not tell you now",
                "Most likely",
                "My sources say no",
                "Outlook good",
                "Cannot predict now",
                "Yes",
                "Outlook not so good",
                "Signs point to yes",
                "Concentrate and ask again",
                "Very doubtful"),
            "Ask, and it will be answered."));

// =ROCKPAPERSCISSORS("Rock") -> "You threw ROCK. Excel threw SCISSORS. You win: Rock smashes Scissors."
ROCKPAPERSCISSORS =
    LAMBDA(
        throw,
        LET(
            _human, UPPER(throw),
            _robot, CHOOSE(RANDBETWEEN(1, 3), "ROCK", "PAPER", "SCISSORS"),
            CONCAT(
                FORMAT("You threw {1}. Excel threw {2}. ", _human, _robot),
                IFS(
                    EQUAL(_human, _robot),
                        FORMAT("Tie: {1} vs. {1}", _human),
                    AND(EQUAL(_human, "PAPER"), EQUAL(_robot, "ROCK")),
                        "You win: Paper covers rock.",
                    AND(EQUAL(_human, "PAPER"), EQUAL(_robot, "SCISSORS")),
                        "You lose: Scissors cuts paper.",
                    AND(EQUAL(_human, "ROCK"), EQUAL(_robot, "PAPER")),
                        "You lose: Paper covers rock.",
                    AND(EQUAL(_human, "ROCK"), EQUAL(_robot, "SCISSORS")),
                        "You win: Rock smashes scissors.",
                    AND(EQUAL(_human, "SCISSORS"), EQUAL(_robot, "ROCK")),
                        "You lose: Rock smashes scissors.",
                    AND(EQUAL(_human, "SCISSORS"), EQUAL(_robot, "PAPER")),
                        "You win: Scissors cuts paper.",
                    else,
                        "Invalid throw. Please choose Rock, Paper, or Scissors"))));

///
/// An Epilogue of Miscellany
///

// On the Uncommon Conception of Readability, or A Brief Dissertation Regarding Aesthetic Principia
ALIAS = // A well-pondered alias deserves its own line.
    LAMBDA( // LAMBDA, our liberator, also deserves its own line.
        parameter, [optional], // Inputs share a single line, unless there be egregiously(?) many.
        LET( // The output clause of LAMBDA on the line following inputs.
            first, ID(parameter), // LET name-value pairings share a line.
            _range, {1,2,3,4,5;6,7,8,9,0}, // Helper aliases are encouraged to reduce line lengths.
            result, INDEX(_range, 1, 1), // Adequately-concise Excel functions can fit on one line.
            IF( // The output clause of LET, like LAMBDA, occupies the line following name-value pairs.
                EQUAL(2, INCREMENT(result)), // IF condition on its own line, preferred over one-liner.
                "Cheers!", // The success case occupies its own line.
                "Call Bertrand."))); // Terminating parentheses live on the last line. (Non-negotiable)
