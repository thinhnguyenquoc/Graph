CREATE PROCEDURE [dbo].[GetProductivity]
	@param1 datetime ,
	@param2 datetime
AS
	select distinct p.Code, p.Name, p.Note, p.Category, p.Duration, p.Price 
	from Schedules s join Programs p on s.Code = p.Code
	where s.Date between @param1 and @param2

RETURN 

--exec  GetProductivity '1/1/2015', '1/1/2015 23:59:59'