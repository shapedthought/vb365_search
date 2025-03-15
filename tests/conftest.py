"""
Common pytest fixtures for vb365_search tests
"""

import pytest

from vb365_search.authentication.auth_models import AuthHeaders
from vb365_search.models.models import Configuration


@pytest.fixture
def mock_config(mocker):
    """
    Fixture for a mock Configuration object
    """
    config = mocker.Mock(spec=Configuration)
    config.vb365 = mocker.Mock()
    config.vb365.api_address = "test-server"
    config.vb365.version = "v8"
    config.microsoft = mocker.Mock()
    config.microsoft.tenant_name = "test-tenant"
    config.microsoft.application_id = "test-app-id"
    return config


@pytest.fixture
def mock_auth_headers(mocker):
    """
    Fixture for mock authentication headers
    """
    return mocker.Mock(spec=AuthHeaders)


@pytest.fixture
def mock_restore_session_id():
    """
    Fixture for a mock restore session ID
    """
    return "test-session-id"
