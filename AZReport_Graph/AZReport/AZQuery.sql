select s.code, Count(s.code) , p.Duration
from Schedules s join Programs p on s.Code = p.Code
where s.Date between '1/1/2015' and '1/1/2015 23:59:59'
group by s.Code, p.Duration

