def _ukedag_eller_helligdag(self, dag_nummer):
        # Convert day number to a date object
        year = datetime.datetime.now().year  # Get the current year
        dato = datetime.datetime(year, 1, 1) + datetime.timedelta(dag_nummer - 1)

        # Check if the date is a weekend (Saturday or Sunday)
        if dato.weekday() >= 5:
            return "helg"

        # Check if the date is a holiday in the Norwegian calendar
        norske_helligdager = [
            (1, 1),  # New Year's Day
            (5, 1),  # Labor Day
            (5, 17),  # Constitution Day
            (12, 25),  # Christmas Day
            (12, 26),  # Boxing Day
        ]

        if (dato.month, dato.day) in norske_helligdager:
            return "helg"

        # If neither weekend nor holiday, consider it a regular weekday
        return "ukedag"