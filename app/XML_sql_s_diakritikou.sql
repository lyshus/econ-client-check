SELECT
    cisloPojistence,
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
FROM xml_data ORDER BY jmeno_prijmeni;