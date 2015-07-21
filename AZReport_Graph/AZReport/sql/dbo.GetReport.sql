CREATE PROCEDURE [dbo].[GetReport]
	@param1 datetime ,
	@param2 datetime
AS
	Declare @totaltime Table(
		code varchar(max) not null,
		freq int,
		Duration datetime,
		ScheduleDate datetime,
		Quantity int
	) 
	while @param1 <= @param2
	begin
		insert into @totaltime select s.code, Count(s.code) freq, p.Duration, CAST(s.Date AS DATE), x.Quantity
		from Schedules s join Programs p on s.Code = p.Code join Sales x on x.Code = p.Code
		where CAST(s.Date AS DATE) = @param1 and DATEPART(minute,p.Duration) > 4
		group by s.Code, p.Duration, CAST(s.Date AS DATE), x.Quantity
	set @param1 = DATEADD(day,1,@param1)
	end
	select * from @totaltime;

	--select distinct p.code, p.Name, t.freq, t.Duration, p.Category, p.Price, p.Note, s.Quantity, s.Date from @totaltime t join Programs p on t.code = p.code join Sales s on s.Code = p.Code
	--where s.Date between @param1 and @param2
	--group by s.Date,p.code, p.Name, t.freq, t.Duration, p.Category, p.Price, p.Note, s.Quantity, s.Date



RETURN 

--exec  GetReport '1/1/2015', '1/3/2015 23:59:59'