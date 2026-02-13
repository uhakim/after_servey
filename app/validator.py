import pandas as pd


def build_validation_frame(errors):
    if not errors:
        return pd.DataFrame(columns=["row", "name", "field", "value", "issue"])
    return pd.DataFrame(errors, columns=["row", "name", "field", "value", "issue"])
