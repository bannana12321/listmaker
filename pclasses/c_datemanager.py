import calendar


class datetabler:

  def month_dict(self, month, year):
    """Return a dict listing each day of the month with its weekday and a 'marked' flag."""
    days_in_month = calendar.monthrange(year, month)[1]
    weekday_names = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    ]

    result = {}
    for day in range(1, days_in_month + 1):
      weekday = calendar.weekday(year, month, day)  # 0 = Monday, 6 = Sunday
      result[str(day)] = {
          "weekday": weekday_names[weekday],
          "marked": False,
      }
    return result

  def toggle_single_day(self, dmonth, day_number):
    """Toggle the 'marked' boolean from True to False (and vice versa) for a single day."""
    day_key = str(day_number)
    if day_key in dmonth:
      dmonth[day_key]["marked"] = not dmonth[day_key]["marked"]
    return dmonth

  def toggle_days(self, dmonth, dayoftheweek):
    """Set/toggle 'marked' for every day of a specific weekday in a month_dict result."""
    for day_key, day_data in dmonth.items():
      if day_data["weekday"] == dayoftheweek:
        self.toggle_single_day(dmonth, int(day_key))
    return dmonth

