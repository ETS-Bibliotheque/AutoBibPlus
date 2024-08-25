from typing import Union, Literal
import pandas as pd

from ..superclasses.insLookup import InsLookup
from ..utils.get_content import get_content
from ..utils.parse_content import chained_get
from ..utils.constants import URLS

class InstitutionLookup(InsLookup):
    # Class variables
    yearRange_list = Literal['3yrs', '3yrsAndCurrent', '3yrsAndCurrentAndFuture', '5yrs', '5yrsAndCurrent', '5yrsAndCurrentAndFuture', '10yrs']
    metricType_list = Literal['AcademicCorporateCollaboration', 'AcademicCorporateCollaborationImpact', 'Collaboration', 'CitationCount', 'CitationsPerPublication', 'CollaborationImpact', 'CitedPublications', 'FieldWeightedCitationImpact', 'ScholarlyOutput', 'PublicationsInTopJournalPercentiles', 'OutputsInTopCitationPercentiles']
    includedDocs_list = Literal['AllPublicationTypes', 'ArticlesOnly', 'ArticlesReviews', 'ArticlesReviewsConferencePapers', 'ArticlesReviewsConferencePapersBooksAndBookChapters', 'ConferencePapersOnly', 'ArticlesConferencePapers', 'BooksAndBookChapters']
    journalImpactType_list = Literal['CiteScore', 'SNIP', 'SJR']

    '''@property
    def name(self):
        return chained_get(self._results, ['institution', 'name'], 'Unknown')
    
    @property
    def id(self):
        return chained_get(self._results, ['institution', 'id'])
    
    @property
    def irl(self):
        return chained_get(self._results, ['institution', 'irl'])

    @property
    def dataSource(self):
        return [[key, value] for key, value in self._dataSource.items()]'''

    def __init__(self, 
                  institution_id: Union[int, str], 
                  api_key: str,
                  token: str,
                  refresh: Union[bool, int] = False, 
                  **kwds: str) -> None:
        """Interaction with the Institution Retrieval API.

        :param institution_id: The ID of the institution.
        :param refresh: Whether to refresh the cached file if it exists or not.
                        If int is passed, cached file will be refreshed if the
                        number of days since last modification exceeds that value.
        :param kwds: Keywords passed on as query parameters.
        """
        # Load json
        self._id = str(institution_id)
        self._api_key = str(api_key)
        self._token = str(token)
        self._refresh = refresh

        InsLookup.__init__(self, 
                        api='InstitutionLookup', 
                        identifier=self._id, 
                        complement="metrics", 
                        **kwds)
        self.kwds = kwds

        # Parse json
        '''self._results = self._json['results']
        self._dataSource = self._json['dataSource']'''

    def __str__(self) -> str:
        """Return a summary string."""
        date = self.get_cache_file_mdate().split()[0]
        s = f"Your choice, as of {date}:\n"\
            f"\t- Name: {self.name}\n"\
            f"\t- ID: \t{self.id}"
        return s

    def _get_metrics_rawdata(self,
                             institution_ids: str = '',
                             metricType: metricType_list = 'ScholarlyOutput',
                             yearRange: yearRange_list = '5yrs',
                             subjectAreaFilterURI: str = '',
                             includeSelfCitations: bool = True,
                             byYear: bool = False,
                             includedDocs: includedDocs_list = 'AllPublicationTypes',
                             journalImpactType: journalImpactType_list = 'CiteScore',
                             showAsFieldWeighted: bool = False,
                            ) -> any:

        institution_ids = self._id if institution_ids == '' else institution_ids

        params = {
            "institutionIds": institution_ids,
            "metricTypes": metricType,
            "yearRange": yearRange,
            "subjectAreaFilterURI": subjectAreaFilterURI,
            "includeSelfCitations": includeSelfCitations,
            "byYear": byYear,
            "includedDocs": includedDocs,
            "journalImpactType": journalImpactType,
            "showAsFieldWeighted": showAsFieldWeighted,
            "apiKey":self._api_key,
            "insttoken":self._token,
        }

        response = get_content(url=URLS['InstitutionLookup']+'metrics', api='InstitutionLookup', params=params, **self.kwds)
        data = response.json()['results'][0]['metrics'][0]
        # data = response.json()['results'][0]['metrics'][0]
        '''last_key = list(data.keys())[-1]
        return data[last_key]'''
        return data

    def get_metrics_Collaboration(self,
                                  institution_ids: str = '',
                                  metricType: Literal['Institutional collaboration', 'AcademicCorporateCollaborationImpact', 'Collaboration', 'CollaborationImpact'] = 'Collaboration',
                                  # collabType: Literal['Academic-corporate collaboration', 'No academic-corporate collaboration', 'Institutional collaboration', 'International collaboration', 'National collaboration', 'Single authorship'] = 'No academic-corporate collaboration',
                                  collabType:Literal['Institutional collaboration', 'AcademicCorporateCollaborationImpact', 'Collaboration', 'CollaborationImpact'] = 'Institutional collaboration',
                                  value_or_percentage: Literal['valueByYear', 'percentageByYear'] = 'valueByYear',
                                  yearRange: yearRange_list = '5yrs',
                                  subjectAreaFilterURI: str = '',
                                  includeSelfCitations: bool = True,
                                  byYear: bool = False,
                                  includedDocs: includedDocs_list = 'AllPublicationTypes',
                                  journalImpactType: journalImpactType_list = 'CiteScore',
                                  showAsFieldWeighted: bool = False,
                                  ):

        return self._get_metrics_rawdata(institution_ids, metricType, yearRange, subjectAreaFilterURI, 
                                         includeSelfCitations, byYear, includedDocs, journalImpactType, showAsFieldWeighted)