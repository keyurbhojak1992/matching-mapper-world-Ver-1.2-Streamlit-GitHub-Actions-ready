"""Regression tests for Demand Mapper and Advanced Demand fill/status parity."""
import pandas as pd
import pytest

from original_master_code import Filter_Demand_Stock_Process as regular
from original_master_code import Advance_Demand_Filter_Process as advanced


@pytest.mark.parametrize('engine', (regular, advanced))
def test_inclusion_short_and_full_names_and_fill(engine):
    df = pd.DataFrame({
        'Serial No': [1, 2, 3],
        'TINC_NAME': ['NONE','MINOR FEATHER','MEDIUM FEATHER'],
        'SINC_NAME': ['NONE','NONE','NONE'],
    })
    rule = pd.Series({'TINC_NAME':'NONE, MIN FEATHER', 'SINC_NAME':'NONE', 'REMARKS':'INCLUDE'})
    included = engine.apply_tolerance_filter(df, rule, 1)
    assert list(included['Serial No']) == [1, 2, 3]
    labels = engine.classify_match_types_vectorized(included, rule)
    assert labels.tolist() == ['Exact', 'Exact', 'Nearby']


@pytest.mark.parametrize('engine', (regular, advanced))
def test_other_ordered_grade_not_false_green(engine):
    df = pd.DataFrame({'Serial No':[1,2], 'TBINC_NAME':['NONE','PIN POINT BLK']})
    rule = pd.Series({'TBINC_NAME':'NONE', 'REMARKS':'INCLUDE'})
    included = engine.apply_tolerance_filter(df, rule, 1)
    assert set(included['Serial No']) == {1,2}
    labels = engine.classify_match_types_vectorized(included, rule)
    assert labels.tolist() == ['Exact','Nearby']


@pytest.mark.parametrize('engine', (regular, advanced))
def test_sgp_comment_uses_contains_and_negative_terms(engine):
    df = pd.DataFrame({'Serial No':[1,2,3], 'SGPComment':['Exclusive Clarity','With exclusive clarity, pin-point','Brown issue']})
    rule = pd.Series({'SGP COMMENT':'EXCLUSIVE CLARITY,!PIN POINT', 'REMARKS':'INCLUDE'})
    matched = engine.apply_tolerance_filter(df, rule, 1)
    assert matched['Serial No'].tolist() == [1]
    assert engine.classify_match_types_vectorized(matched,rule).tolist() == ['Exact']
