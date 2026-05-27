-- Flag fixed-date Malawi holidays as recurring (match by year-2026 + description).
UPDATE Holidays SET Recurring = 1
WHERE Description IN (
  'New Year Day', 'John Chilembwe Day', 'Martyrs Day',
  'Labour Day', 'Kamuzu Day', 'Freedom Day',
  'Independence Day', 'Christmas Day', 'Boxing Day'
);

-- Delete the 2027 copies of those recurring holidays
-- (the 2026 row now covers every year automatically).
DELETE FROM Holidays
WHERE YEAR(Holiday) = 2027
  AND Description IN (
    'New Year Day', 'John Chilembwe Day', 'Martyrs Day',
    'Labour Day', 'Kamuzu Day', 'Freedom Day',
    'Independence Day', 'Christmas Day', 'Boxing Day'
  );

SELECT Holiday_DPA_, CONVERT(varchar(10), Holiday, 23) AS Holiday, Description, Recurring
FROM Holidays
ORDER BY Recurring DESC, Holiday;
