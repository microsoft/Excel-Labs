// Spreadsheet Lisp (SL) v0.5.0
// October 17, 2023

// SL was designed using the Advanced Formula Editor (AFE) from Microsoft Garage.
// All design decisions codified herein are subject to change following sufficient argumentation.

// On the Uncommon Conception of Readability, or A Brief Dissertation Regarding Aesthetic Principia
ALIAS = // A well-pondered alias deserves its own line.
    LAMBDA( // LAMBDA, our liberator, also deserves its own line.
        parameter, [optional], // Inputs share a single line, unless there be egregiously(?) many
        LET( // Output clause of LAMBDA on separate line after inputs
            first, ID(parameter), // LET name-value pairings share a line
            _range, {1,2,3,4,5;6,7,8,9,0}, // Helper aliases are encouraged to reduce line lengths
            result, INDEX(_range, 1, 1), // Adequately-concise Excel functions can fit on one line
            IF( // Output clause of LET on last line following name-value pairs
                INCREMENT(result)=2, // Condition of IF() on its own line
                "Cheers!", // Followed by the success case on its own line, etc.
                "Call Bertrand."))); // All terminating parentheses on last line. This is non-negotiable.

// Constant Aliases
else = TRUE;
t = TRUE;
f = FALSE;

ID =
    LAMBDA(
        x,
        x);

CAR =
    LAMBDA(
        range,
        INDEX(range, 1, 1));

_FIRST =
    LAMBDA(
        range,
        CAR(range));

_SECOND =
    LAMBDA(
        range,
        CAR(CDR(range)));

_THIRD =
    LAMBDA(
        range,
        CAR(CDR(CDR(range))));

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

CONS =
    LAMBDA(
        head, body,
        LET(
            horizontal, COLUMNS(body)>ROWS(body),
            IF(
                horizontal,
                HSTACK(head, body),
                VSTACK(head, body))));

EQ =
    LAMBDA(
        range,
        AND(EXACT(CAR(range), range)));

EQUAL =
    LAMBDA(
        a, b,
        a=b);

/*
This function requires macros to be enabled.
EVAL =
    LAMBDA(
        expression_string,
        EVALUATE(expression_string));
*/

IS =
    LAMBDA(
        argument,
        IF(
            ISOMITTED(argument),
            0,
            1));

VLIST =
    LAMBDA(
        [a], [b], [c], [d], [e], [f], [g], [h], [i], [j],
        LET(
            row_count,SUM(IS(a), IS(b), IS(c), IS(d), IS(e), IS(f), IS(g), IS(h), IS(i), IS(j)),
            MAKEARRAY(
                row_count,
                1,
                LAMBDA(
                    row,
                    _,
                    CHOOSE(row, a, b, c, d, e, f, g, h, i, j)))));

HLIST =
    LAMBDA(
        [a], [b], [c], [d], [e], [f], [g], [h], [i], [j],
        LET(
            col_count,SUM(IS(a), IS(b), IS(c), IS(d), IS(e), IS(f), IS(g), IS(h), IS(i), IS(j)),
            MAKEARRAY(
                1,
                col_count,
                LAMBDA(
                    _,
                    col,
                    CHOOSE(col, a, b, c, d, e, f, g, h, i, j)))));

MEMBER =
    LAMBDA(
        needle, haystack,
        OR(EXACT(needle, haystack)));

ATOM =
    LAMBDA(
        expression,
        LET(
            number, 1,
            text, 2,
            boolean, 4,
            atoms, HLIST(number, text, boolean),
            MEMBER(TYPE(expression), atoms)));

FIRSTROW =
    LAMBDA(
        range,
        INDEX(range, 1, 0));

FIRSTCOL =
    LAMBDA(
        range,
        INDEX(range, 0, 1));

NEGATE =
    LAMBDA(
        number,
        PRODUCT(number, -1));

DECREMENT =
    LAMBDA(
        number,
        SUM(number, NEGATE(1)));

INCREMENT =
    LAMBDA(
        number,
        SUM(number, 1));

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

HOMOGENOUS =
    LAMBDA(
        range,
        EQ(range));

HETEROGENOUS =
    LAMBDA(
        range,
        NOT(HOMOGENOUS(range)));

PROVIDED =
    LAMBDA(
        variable,
        NOT(ISOMITTED(variable)));

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

GET =
    LAMBDA(
        range, [sheet], [file], [path],
        LET(
            sheet_provided, PROVIDED(sheet),
            file_provided, PROVIDED(file),
            path_provided, PROVIDED(path),
            reference_wrapper, IF(NOT(ISERROR(FIND(" ", sheet))), "'", ""),
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

FORMAT =
    LAMBDA(
        template, [first], [second], [third], [fourth], [fifth],
        LET(
            _after1, IF(PROVIDED(first), SUBSTITUTE(template, "{1}", first), template),
            _after2, IF(PROVIDED(second), SUBSTITUTE(_after1, "{2}", second), _after1),
            _after3, IF(PROVIDED(third), SUBSTITUTE(_after2, "{3}", third), _after2),
            _after4, IF(PROVIDED(fourth), SUBSTITUTE(_after3, "{4}", fourth), _after3),
            IF(PROVIDED(fifth), SUBSTITUTE(_after4, "{5}", fifth), _after4)));

// Currently supports single-condition where clause.
// =SQL(range, "select [col], [etc.] where [col] >= 23")
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
GETCOL =
    LAMBDA(
        index, range,
        INDEX(range, 0, index));

GETCOLS =
    LAMBDA(
        range, indexes,
        IF(
            COLUMNS(indexes) = 1,
            GETCOL(CAR(indexes), range),
            HSTACK(
                GETCOL(CAR(indexes), range),
                GETCOLS(range, CDR(indexes)))));

FIRSTWORD =
    LAMBDA(
        sentence,
        LEFT(sentence, TRIM(FIND(" ", sentence)-1)));

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

TRIMALL =
    LAMBDA(
        list,
        MAP(list, LAMBDA(item, TRIM(item))));

GREATERTHAN =
    LAMBDA(
        x, y,
        x>y);

GTE =
    LAMBDA(
        x, y,
        x>=y);

LESSTHAN =
    LAMBDA(
        x, y,
        x<y);

LTE =
    LAMBDA(
        x, y,
        x<=y);

ORIENTATION =
    LAMBDA(
        range,
        LET(
            horizontal, GREATERTHAN(COLUMNS(range), ROWS(range)),
            IF(
                horizontal,
                "Horizontal",
                "Vertical")));
                    
CAGR =
    LAMBDA(
        beginning_period, ending_period, [periods],
        LET(
            p, IF(PROVIDED(periods), periods, 1),
            nominal_growth, ending_period/beginning_period,
            POWER(nominal_growth, 1/p)-1));

FACTORIAL =
    LAMBDA(
        n,
        IF(
            EQUAL(n, 1),
            1,
            PRODUCT(n, FACTORIAL(DECREMENT(n)))));
