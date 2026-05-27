DECLARE @next int = (SELECT ISNULL(MAX(Holiday_DPA_),0)+1 FROM Holidays);
WITH H(d, n) AS (SELECT * FROM (VALUES
  -- 2026
  (CAST('2026-01-01' AS date), 'New Year Day'),
  ('2026-01-15',               'John Chilembwe Day'),
  ('2026-03-03',               'Martyrs Day'),
  ('2026-04-03',               'Good Friday'),
  ('2026-04-06',               'Easter Monday'),
  ('2026-05-01',               'Labour Day'),
  ('2026-05-14',               'Kamuzu Day'),
  ('2026-06-14',               'Freedom Day'),
  ('2026-07-06',               'Independence Day'),
  ('2026-10-12',               'Mothers Day'),
  ('2026-12-26',               'Boxing Day'),
  -- 2027
  ('2027-01-01',               'New Year Day'),
  ('2027-01-15',               'John Chilembwe Day'),
  ('2027-03-03',               'Martyrs Day'),
  ('2027-03-26',               'Good Friday'),
  ('2027-03-29',               'Easter Monday'),
  ('2027-05-01',               'Labour Day'),
  ('2027-05-14',               'Kamuzu Day'),
  ('2027-06-14',               'Freedom Day'),
  ('2027-07-06',               'Independence Day'),
  ('2027-10-11',               'Mothers Day'),
  ('2027-12-25',               'Christmas Day'),
  ('2027-12-26',               'Boxing Day')
) AS v(d,n))
INSERT INTO Holidays (Holiday_DPA_, Holiday, Description)
SELECT @next + ROW_NUMBER() OVER (ORDER BY d) - 1, d, n
FROM H WHERE NOT EXISTS (SELECT 1 FROM Holidays x WHERE x.Holiday = H.d);
SELECT @@ROWCOUNT AS Inserted;
SELECT Holiday_DPA_, CONVERT(varchar(10), Holiday, 23) AS Holiday, Description
FROM Holidays ORDER BY Holiday;
