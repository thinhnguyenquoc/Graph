CREATE PROCEDURE [dbo].[GetTotalTime]
	@param1 datetime ,
	@param2 datetime
AS
	Declare @totaltime Table(
		code varchar(max) not null,
		freq int,
		Duration datetime
	) 
	insert into @totaltime select s.code, Count(s.code) freq, p.Duration
	from Schedules s join Programs p on s.Code = p.Code
	where s.Date between @param1 and @param2
	group by s.Code, p.Duration

	select distinct p.code, p.Name, t.freq, t.Duration, p.Category, p.Price, p.Note, s.Quantity, s.Date from @totaltime t join Programs p on t.code = p.code join Sales s on s.Code = p.Code
	where s.Date between @param1 and @param2
	group by s.Date,p.code, p.Name, t.freq, t.Duration, p.Category, p.Price, p.Note, s.Quantity



RETURN 

--exec  GetTotalTime '3/1/2015', '3/1/2015 23:59:59'