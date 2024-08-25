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

    @property
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
        return [[key, value] for key, value in self._dataSource.items()]

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
        self._results = self._json['results']
        self._dataSource = self._json['dataSource']

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
                             byYear: bool = True,
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
        data = response.json()['results']['metrics']
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
                                  byYear: bool = True,
                                  includedDocs: includedDocs_list = 'AllPublicationTypes',
                                  journalImpactType: journalImpactType_list = 'CiteScore',
                                  showAsFieldWeighted: bool = False,
                                  ):
        '''if metricType in ('AcademicCorporateCollaboration', 'AcademicCorporateCollaborationImpact'):
            self._check_args(collabType, metricType, ('Academic-corporate collaboration', 'No academic-corporate collaboration'))'''
        if metricType in ('Collaboration', 'CollaborationImpact'):
            self._check_args(collabType, metricType, ('Institutional collaboration', 'International collaboration', 'National collaboration', 'Single authorship'))
        if metricType in ('AcademicCorporateCollaborationImpact', 'CollaborationImpact'):
            value_or_percentage = 'valueByYear'

        return MetricsFormatage(self._for_advanced_metrics(metricType, yearRange, subjectAreaFilterURI, includeSelfCitations, byYear, includedDocs, journalImpactType, showAsFieldWeighted,
                                                           institution_ids, 'collabType', collabType, value_or_percentage))

    def get_metrics_Percentile(self,
                               institution_ids: str = '',
                               metricType: Literal['PublicationsInTopJournalPercentiles', 'OutputsInTopCitationPercentiles'] = 'OutputsInTopCitationPercentiles',
                               threshold: Literal[1, 5, 10, 25] = 10,
                               value_or_percentage: Literal['valueByYear', 'percentageByYear'] = 'valueByYear',
                               yearRange: yearRange_list = '5yrs',
                               subjectAreaFilterURI: str = '',
                               includeSelfCitations: bool = True,
                               byYear: bool = True,
                               includedDocs: includedDocs_list = 'AllPublicationTypes',
                               journalImpactType: journalImpactType_list = 'CiteScore',
                               showAsFieldWeighted: bool = False,
                               ):
        return MetricsFormatage(self._for_advanced_metrics(metricType, yearRange, subjectAreaFilterURI, includeSelfCitations, byYear, includedDocs, journalImpactType, showAsFieldWeighted,
                                                           institution_ids, 'threshold', threshold, value_or_percentage))

    def get_metrics_Other(self,
                          institution_ids: str = '',
                          metricType: Literal['CitationCount', 'CitationsPerPublication', 'CitedPublications', 'FieldWeightedCitationImpact', 'ScholarlyOutput'] = 'ScholarlyOutput',
                          yearRange: yearRange_list = '5yrs',
                          subjectAreaFilterURI: str = '',
                          includeSelfCitations: bool = True,
                          byYear: bool = True,
                          includedDocs: includedDocs_list = 'AllPublicationTypes',
                          journalImpactType: journalImpactType_list = 'CiteScore',
                          showAsFieldWeighted: bool = False,
                          ):
        return MetricsFormatage({'valueByYear': self._for_advanced_metrics(metricType, yearRange, subjectAreaFilterURI, includeSelfCitations, byYear, includedDocs, journalImpactType, showAsFieldWeighted,
                                                                          institution_ids)})

    def _get_institution_rawdata(self,
                                 institutionId: Union[str, int],
                                 yearRange: yearRange_list = '5yrs',
                                 limit: int = 500,
                                 offset: int = 0):
        
        limit = 1 if limit < 1 else (500 if limit > 500 else limit)

        params = {
            "yearRange": yearRange,
            "limit": limit,
            "offset": offset
        }

        response = get_content(url=URLS['InstitutionLookup']+'institutionId/'+str(institutionId), api='InstitutionLookup', params=params, **self.kwds)
        data = response.json()
        del data['link']
        return data

    def institutional_authors(self, institutionId: Union[str, int], yearRange: yearRange_list = '5yrs'):
        data_list = self._get_institution_rawdata(institutionId=institutionId, yearRange=yearRange, limit=500, offset=0)['authors']
        offset = 500
        while len(data_list) == offset:
            temp = self._get_institution_rawdata(institutionId=institutionId, yearRange=yearRange, limit=500, offset=offset)['authors']
            data_list += temp
            offset += 500
        return pd.DataFrame(data_list)

    '''def get_metrics_DocType(self, institution_ids: str = '', metricType: metricType_list = 'ScholarlyOutput', yearRange: yearRange_list = '5yrs', subjectAreaFilterURI: str = '', includeSelfCitations: bool = True, byYear: bool = True, includedDocs: includedDocs_list = 'AllPublicationTypes', showAsFieldWeighted: bool = False):
        docType = ['ArticlesOnly', 'ConferencePapersOnly', 'ArticlesReviews', 'ArticlesReviewsConferencePapers', 'ArticlesReviewsConferencePapersBooksAndBookChapters', 'BooksAndBookChapters']
        doc_df = pd.DataFrame({'docType': docType})
        doc_df['docTypeData'] = doc_df.apply(lambda x: self._get_metrics_rawdata(institution_ids, metricType, yearRange, subjectAreaFilterURI, includeSelfCitations, byYear, x['docType'], journalImpactType, showAsFieldWeighted), axis=1)
        return doc_df'''

    def _check_args(self, arg, metricType, valid_args):
        assert arg in valid_args, f'{metricType} requires collabType to be {valid_args}, got "{arg}".'
    
    def _for_advanced_metrics(self, metricType, yearRange, subjectAreaFilterURI, includeSelfCitations, byYear, includedDocs, journalImpactType, showAsFieldWeighted, institution_ids, special_key=None, special_value=None, value_or_percentage=None):
        raw = self._get_metrics_rawdata(institution_ids, metricType, yearRange, subjectAreaFilterURI, includeSelfCitations, byYear, includedDocs, journalImpactType, showAsFieldWeighted)
        return raw if special_key is None else {value_or_percentage: {special_value: raw[special_key][special_value][value_or_percentage]}}
    
class MetricsFormatage():
    @property
    def Raw(self):
        return self._dict
        
    @property
    def Dictionary(self):
        return self._dict[list(self._dict.keys())[-1]]

    @property
    def DataFrame(self):
        data_type = list(self._dict.keys())[-1]
        column_name = 'Percentage' if data_type == 'percentageByYear' else 'Value'
        df = pd.DataFrame.from_dict(self._dict[data_type], orient='index', columns=[column_name])
        df.index.name = 'Year'
        return df

    @property
    def List(self):
        data_type = list(self._dict.keys())[-1]
        return [[int(cle) for cle in self._dict[data_type].keys()], list(self._dict[data_type].values())]
    

    def __init__(self, data: dict) -> None:
        self._dict = data

    def __repr__(self):
        return str("You must select a property: 'Raw', 'List', 'Dictionary', 'DataFrame' after calling the InstitutionLookup class method!")    
    

class InstitutionalFormatage():  
    @property
    def Dictionary(self):
        return self._dict
    
    @property
    def List(self):
        data_list = list(self._dict.values())
        names = [item.get('name', '') for item in data_list]
        ids = [item.get('id', '') for item in data_list]
        scholarly_outputs = [item.get('scholarlyOutput', '') for item in data_list]
        return [ids, names, scholarly_outputs]
    
    @property
    def DataFrame(self):
        df = pd.DataFrame.from_dict(self._dict, orient='index')
        df.index.name = 'ID'
        return df


    def __init__(self, data: dict) -> None:
        self._dict = data

    def __repr__(self):
        return str("You must select a property: 'List', 'Dictionary', 'DataFrame' after calling the InstitutionLookup class method!")
    

