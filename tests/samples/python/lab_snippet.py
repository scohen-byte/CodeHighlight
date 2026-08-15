class GradeBook:
    """Track student scores for one course."""
    PASS_MARK = 60

    def record(self, student, score):
        # Scores outside 0-100 are almost always a typo.
        if not 0 <= score <= 100:
            raise ValueError(f"bad {student}: {score}")
        self.scores[student] = score

    def passing(self):
        return [n for n, s in self.scores.items() if s >= 60]
