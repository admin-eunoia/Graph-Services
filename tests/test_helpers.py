from types import SimpleNamespace
import pathlib
import sys

import pytest

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.append(str(PROJECT_ROOT))

from routes.routes import (
    _build_dest_file_name,
    _validate_data_dict,
    _validate_naming_dict,
)
from Services.graph_services import validate_cell_map_or_raise


def _make_template(pattern: str = "{template_key}_{client_name_sanitized}.xlsx"):
    return SimpleNamespace(
        dest_file_pattern=pattern,
        template_key="waman_prueba",
    )


def test_build_dest_file_name_with_custom_pattern():
    template = _make_template("{empresa}_{persona}.xlsx")
    body = {
        "client_key": "eunoia",
        "client_name": "Cliente Demo",
        "naming": {"empresa": "EUNOIA", "persona": "Dan"},
    }

    result = _build_dest_file_name(template, body)

    assert result == "EUNOIA_Dan.xlsx"


def test_build_dest_file_name_sanitizes_client_name():
    template = _make_template()
    body = {
        "client_key": "eunoia",
        "client_name": "../Cliente Demo",
        "naming": {},
    }

    with pytest.raises(ValueError) as exc:
        _build_dest_file_name(template, body)

    assert "Nombre de archivo resultante inválido" in str(exc.value)


def test_build_dest_file_name_missing_placeholder_raises():
    template = _make_template("{empresa}_{persona}.xlsx")
    body = {
        "client_key": "eunoia",
        "client_name": "Cliente Demo",
        "naming": {"empresa": "EUNOIA"},
    }

    with pytest.raises(ValueError) as exc:
        _build_dest_file_name(template, body)

    assert "Faltan campos" in str(exc.value)


@pytest.mark.parametrize(
    "payload,expected_error",
    [
        ({}, "data no puede estar vacío"),
        ({"A1": "x" * 3000}, "Valor de 'A1' excede"),
        ({f"A{i}": i for i in range(501)}, "data excede el máximo"),
    ],
)
def test_validate_data_dict_errors(payload, expected_error):
    error = _validate_data_dict(payload)
    assert error and expected_error in error


def test_validate_data_dict_success():
    payload = {"A1": "valor", "B2": 3}
    assert _validate_data_dict(payload) is None


def test_validate_naming_dict_sanitizes_strings():
    err, result = _validate_naming_dict({"empresa": "  EUNOIA ", "persona": "Dan"})
    assert err is None
    assert result == {"empresa": "EUNOIA", "persona": "Dan"}


@pytest.mark.parametrize(
    "naming,expected",
    [
        ("texto", "naming debe ser un objeto"),
        ({"": "valor"}, "claves inválidas"),
        ({"empresa": "x" * 600}, "excede"),
        ({"invalid key": "valor"}, "no permitidos"),
    ],
)
def test_validate_naming_dict_errors(naming, expected):
    err, _ = _validate_naming_dict(naming)
    assert err and expected in err


def test_validate_cell_map_or_raise_success():
    validate_cell_map_or_raise({"A1": "valor", "Hoja1!B2": 10})


@pytest.mark.parametrize(
    "payload",
    [
        {"A0": "valor"},
        {"Hoja!": 5},
        {"A1": {"complex": "value"}},
    ],
)
def test_validate_cell_map_or_raise_errors(payload):
    with pytest.raises(ValueError):
        validate_cell_map_or_raise(payload)
