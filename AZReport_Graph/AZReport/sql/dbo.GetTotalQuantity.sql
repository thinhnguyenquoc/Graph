CREATE PROCEDURE [dbo].[GetTotalQuantity]
	@param1 datetime ,
	@param2 datetime
AS
	Declare @program Table(
		code varchar(max) not null
	) 

	insert into @program select distinct p.Code 
	from Schedules s join Programs p on s.Code = p.Code
	where s.Date between @param1 and @param2

	--select * from @program

	Declare @quantity Table(
		code varchar(max) not null,
		quantity int
	)

	insert into @quantity select s.Code, s.Quantity
	--from Sales s join @program p on s.Code = p.code 
	from Sales s
	where s.Date between @param1 and @param2 

	select q.code as Code, Sum(q.quantity) as Quantity from @quantity q
	group by q.code

RETURN 0

--exec  GetTotalQuantity '1/3/2015', '1/3/2015 23:59:59'