SELECT substr('0000000000' || rod_cislo, -10) AS rod_cislo,
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
FROM excel_data
WHERE NOT (name_1 = 6 AND celk_odv = 0)
  AND name_1 IS NOT NULL
  AND name_1 != ''
  AND NOT (
    poc_cisel = 0
    AND os_cislo IN (
      SELECT os_cislo
      FROM excel_data
      WHERE poc_cisel = 0
      GROUP BY os_cislo
      HAVING COUNT(*) > 1
    )
  )
ORDER BY jmeno_prijmeni;