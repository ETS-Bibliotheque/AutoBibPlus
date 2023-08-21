"""Superclass to access all SciVal lookup APIs and dump the results."""

from typing import Union, Literal

from .base import Base
from .constants import URLS
from .get_content import get_folder


class Lookup(Base):
    def __init__(self,
                 api: Literal["AuthorLookup", "CountryLookup", "CountryGroupLookup", "InstitutionLookup", 
                              "InstitutionGroupLookup", "PublicationLookup", "ScopusSourceLookup", 
                              "SubjectAreaLookup", "TopicLookup", "TopicClusterLookup", "WorldLookup"],
                 identifier: Union[int, str] = None,
                 complement: str = "",
                 **kwds: str
                 ) -> None:
        """Class intended as superclass to perform retrievals.

        :param api: The name of the Scopus API to be accessed.  Allowed values:
                    AuthorLookup, CountryLookup, CountryGroupLookup,
                    InstitutionLookup, InstitutionGroupLookup, PublicationLookup,
                    ScopusSourceLookup, SubjectAreaLookup, TopicLookup,
                    TopicClusterLookup, WorldLookup.
        :param identifier: The ID to look for.
        :param complement: The URL complement that launches the correct getter 
            from the selected Lookup API.
        :param kwds: Keywords passed on to requests header.  Must contain
                     fields and values specified in the respective
                     API specification.

        Raises
        ------
        KeyError
            If parameter `api` is not one of the allowed values.
        """
        # Construct URL and cache file name
        url = URLS[api] + complement
        if identifier != None:
            stem = identifier.replace('/', '_')

        self._cache_file_path = get_folder(api, None)/stem

        # Parse file contents
        params = {'authors': str(identifier),
                  'metricTypes': "ScholarlyOutput",
                  **kwds}
        Base.__init__(self, params=params, url=url, api=api)
