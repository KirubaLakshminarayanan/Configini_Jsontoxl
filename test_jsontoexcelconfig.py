import datetime
import json
import os
import warnings

import pandas as pd
import pytest

from jsontoexcelconfig import flatten_json, convert_to_excel

pytestmark = pytest.mark.filterwarnings("ignore::DeprecationWarning")  # Ignore deprecation warnings globally



@pytest.fixture(autouse=True)
def suppress_warnings():
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=DeprecationWarning)
        yield


@pytest.fixture
def config_file(tmpdir):
    config_path = tmpdir.join('config.ini')
    config_content = """
    [Files]
    input_dir = {input_dir}
    output_dir = {output_dir}
    chunk = 10

    [Excel]
    sheet_name = Sheet1
    """.format(input_dir=tmpdir.join('input'), output_dir=tmpdir.join('output'))
    with open(config_path, 'w') as f:
        f.write(config_content)
    return config_path


@pytest.fixture
def json_file(tmpdir):
    json_path = tmpdir.join('input/test.json')
    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    data = {
        "name": "John Doe",
        "age": 30,
        "city": "New York"
    }
    with open(json_path, 'w') as f:
        json.dump(data, f)
    return json_path


def convert_to_excel(flattened_data, output_file, sheet_name='Sheet1'):
    df = pd.DataFrame([flattened_data])
    with pd.ExcelWriter(output_file) as writer:
        # Replace occurrences of datetime.datetime.utcnow() with datetime.datetime.now(datetime.timezone.utc)
        now_utc = datetime.datetime.now(datetime.timezone.utc)
        # Add a sheet and make it active
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        # Ensure the created sheet is set as active and visible
        #workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        workbook.active = workbook.index(worksheet)

        # Fixing the datetime usage
        workbook.properties.modified = datetime.datetime.now(datetime.timezone.utc)
        workbook.properties.created = datetime.datetime.now(datetime.timezone.utc)


def test_flatten_json_empty():
    data = {}
    assert flatten_json(data) == {}


def test_flatten_json_valid():
    data = {"name": "John", "age": 30}
    assert flatten_json(data) == {"name": "John", "age": 30}



def test_flatten_json_invalid():
    data = "invalid_json"
    assert flatten_json(data) == {}


def test_flatten_json_nested():
    data = {"person": {"name": "John", "age": 30}}
    assert flatten_json(data) == {"person_name": "John", "person_age": 30}


def test_flatten_json_list():
    data = {"person": [{"name": "John", "age": 30}, {"name": "Alice", "age": 25}]}
    assert flatten_json(data) == {
        "person_0_name": "John", "person_0_age": 30,
        "person_1_name": "Alice", "person_1_age": 25
    }



def test_flatten_json_individual():
    data = {"name": "John", "age": 30, "city": "New York"}
    assert flatten_json(data) == {"name": "John", "age": 30, "city": "New York"}


def test_flatten_json_batch():
    data = {
        "persons": [
            {"name": "John", "age": 30, "city": "New York"},
            {"name": "Alice", "age": 25, "city": "Los Angeles"},
            {"name": "Bob", "age": 35, "city": "Chicago"}
        ]
    }
    flattened_data = flatten_json(data)
    assert flattened_data == {
        "persons_0_name": "John", "persons_0_age": 30, "persons_0_city": "New York",
        "persons_1_name": "Alice", "persons_1_age": 25, "persons_1_city": "Los Angeles",
        "persons_2_name": "Bob", "persons_2_age": 35, "persons_2_city": "Chicago"
    }


def test_flatten_json_duplicate_keys():
    data = {"name": "John", "name": "Doe"}
    assert flatten_json(data) == {"name": "Doe"}


def test_flatten_json_nested_list():
    data = {"person": [{"name": "John", "age": 30}, {"name": "Alice", "age": 25}], "city": ["New York", "Los Angeles"]}
    assert flatten_json(data) == {
        "person_0_name": "John", "person_0_age": 30,
        "person_1_name": "Alice", "person_1_age": 25,
        "city_0": "New York", "city_1": "Los Angeles"
    }


def test_flatten_json_complex_nested():
    data = {"person": {"name": "John", "address": {"city": "New York", "country": "USA"}}}
    assert flatten_json(data) == {
        "person_name": "John", "person_address_city": "New York",
        "person_address_country": "USA"
    }


def test_flatten_json_null_value():
    data = {"name": "John", "age": None}
    assert flatten_json(data) == {"name": "John", "age": None}


def test_flatten_json_boolean_value():
    data = {"name": "John", "is_adult": True}
    assert flatten_json(data) == {"name": "John", "is_adult": True}


@pytest.mark.filterwarnings("ignore::DeprecationWarning")
def test_convert_to_excel(tmpdir):
    data = {"name": "John", "age": 30, "city": "New York"}
    flattened_data = flatten_json(data)
    output_file = str(tmpdir.join('output.xlsx'))
    convert_to_excel(flattened_data, output_file, sheet_name='Sheet1')

    # Read the Excel file to verify its content
    df = pd.read_excel(output_file)
    expected_data = {
        "name": ["John"],
        "age": [30],
        "city": ["New York"]
    }
    expected_df = pd.DataFrame(expected_data)
    pd.testing.assert_frame_equal(df, expected_df)