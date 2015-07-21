CREATE PROCEDURE [dbo].[GetTotalFrequency]
	@param1 datetime ,
	@param2 datetime
AS
	Declare @temp Table(
		Code varchar(max) not null,
		Date date
	) 
	insert into @temp select s.Code, CONVERT(date, s.Date) 
	from Schedules s join Programs p on s.Code = p.Code
	where s.Date between @param1 and @param2

	select t.Code, t.Date, COUNT(t.Code) Freq
	from @temp t
	group by t.Code, t.Date
RETURN 

--exec  GetTotalFrequency '1/1/2015', '1/3/2015 23:59:59'