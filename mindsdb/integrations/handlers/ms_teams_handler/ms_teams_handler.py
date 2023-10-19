import msal
import requests
from typing import Optional

from mindsdb.integrations.handlers.ms_teams_handler.ms_teams_tables import MessagesTable
from mindsdb.integrations.libs.api_handler import APIHandler
from mindsdb.integrations.libs.response import (
    HandlerStatusResponse as StatusResponse,
)

from mindsdb.utilities import log
from mindsdb_sql import parse_sql


class MSTeamsHandler(APIHandler):
    """
    The Microsoft Teams handler implementation.
    """

    MICROSOFT_GRAPH_BASE_API_URL: str = "https://graph.microsoft.com/"
    MICROSOFT_GRAPH_API_VERSION: str = "v1.0"
    MICROSOFT_GRAPH_API_SCOPES: list = ["https://graph.microsoft.com/.default"]

    name = 'teams'

    def __init__(self, name: str, **kwargs):
        """
        Initialize the handler.
        Args:
            name (str): name of particular handler instance
            **kwargs: arbitrary keyword arguments.
        """
        super().__init__(name)

        connection_data = kwargs.get("connection_data", {})
        self.connection_data = connection_data
        self.kwargs = kwargs

        self.connection = None
        self.is_connected = False

        messages_data = MessagesTable(self)
        self._register_table("messages", messages_data)

    def connect(self):
        """
        Set up the connection required by the handler.
        Returns
        -------
        StatusResponse
            connection object
        """
        if self.is_connected is True:
            return self.connection

        self.connection = msal.ConfidentialClientApplication(
            client_id=self.connection_data.get('client_id'),
            client_credential=self.connection_data.get('client_secret'),
            authority=f"https://login.microsoftonline.com/" f"{self.connection_data['tenant_id']}",
        )

        self.is_connected = True

        return self.connection
    
    def _get_access_token(self):
        """
        Get the API token.
        Returns
        -------
        str
            API token
        """
        token = self.connection.acquire_token_silent(
            scopes=self.MICROSOFT_GRAPH_API_SCOPES,
            account=None,
        )

        if not token:
            token = self.connection.acquire_token_for_client(scopes=self.MICROSOFT_GRAPH_API_SCOPES)

        return token

    def check_connection(self) -> StatusResponse:
        """
        Check connection to the handler.
        Returns:
            HandlerStatusResponse
        """

        response = StatusResponse(False)

        try:
            self.connect()
            self._get_access_token()
            response.success = True
        except Exception as e:
            log.logger.error(f'Error connecting to Microsoft Teams!')
            response.error_message = str(e)

        self.is_connected = response.success

        return response
    
    def call_graph_api(self, api_endpoint: str, method: str = "GET", data: Optional[dict] = None) -> dict:
        api_url = f"{self.MICROSOFT_GRAPH_BASE_API_URL}{self.MICROSOFT_GRAPH_API_VERSION}/{api_endpoint}/"
        headers = {
            "Authorization": f"Bearer {self._get_access_token()['access_token']}",
            "Content-Type": "application/json",
        }

        if method == "GET":
            response = requests.get(api_url, headers=headers)
        elif method == "POST":
            response = requests.post(api_url, headers=headers, json=data)
        else:
            raise ValueError(f"Unsupported method '{method}'.")
        
        if response.status_code == 200:
            return response.json()
        else:
            raise requests.exceptions.RequestException(response.text)

    def native_query(self, query: str) -> StatusResponse:
        """Receive and process a raw query.
        Parameters
        ----------
        query : str
            query in a native format
        Returns
        -------
        StatusResponse
            Request status
        """
        ast = parse_sql(query, dialect="mindsdb")
        return self.query(ast)