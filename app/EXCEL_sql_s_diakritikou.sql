SELECT
    substr('0000000000' || rod_cislo, -10) AS rod_cislo,
    (
        SELECT GROUP_CONCAT(
            UPPER(SUBSTR(word, 1, 1)) || LOWER(SUBSTR(word, 2)), ' '
        )
        FROM (
            SELECT TRIM(value) AS word
            FROM json_each(
                '["' || REPLACE(
                    REPLACE(
                        REPLACE(
                            REPLACE(jmeno || ' ' || prijmeni, '-', ' '), 
                        '_', ' '), 
                    '  ', ' '), 
                ' ', '","') || '"]'
            )
        )
    ) AS jmeno_prijmeni
FROM (
    SELECT *
    FROM excel_data
    WHERE poc_cisel != 0
      AND NOT (name_1 = 6 AND celk_odv = 0)
      AND name_1 IS NOT NULL
      AND name_1 != ''
    GROUP BY os_cislo
)
ORDER BY jmeno_prijmeni;