import pandas as pd
from typing import Text, List, Dict, Any

from mindsdb_sql.parser import ast
from mindsdb.integrations.libs.api_handler import APITable

from mindsdb.integrations.handlers.utilities.query_utilities.insert_query_utilities import INSERTQueryParser

from mindsdb.utilities.log import get_log

logger = get_log("integrations.ms_teams_handler")


class ChannelsTable(APITable):
    pass


class ChatsTable(APITable):
    pass

