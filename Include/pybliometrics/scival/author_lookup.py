from typing import Union, Literal
import pandas as pd


from ..superclasses.lookup import Lookup
from ..utils.get_content import get_content
from ..utils.parse_content import chained_get
from ..utils.constants import URLS


class AuthorLookup(Lookup):
    # Class variables
    yearRange_list = Literal['3yrs', '3yrsAndCurrent', '3yrsAndCurrentAndFuture', '5yrs', '5yrsAndCurrent', '5yrsAndCurrentAndFuture', '10yrs']
    metricType_liste = Literal['AcademicCorporateCollaboration', 'AcademicCorporateCollaborationImpact', 'Collaboration','CitationCount', 'CitationsPerPublication', 'CollaborationImpact', 'CitedPublications','FieldWeightedCitationImpact', 'ScholarlyOutput', 'PublicationsInTopJournalPercentiles', 'OutputsInTopCitationPercentiles']
    includedDocs_list = Literal['AllPublicationTypes', 'ArticlesOnly', 'ArticlesReviews', 'ArticlesReviewsConferencePapers', 'ArticlesReviewsConferencePapersBooksAndBookChapters', 'ConferencePapersOnly', 'ArticlesConferencePapers', 'BooksAndBookChapters']
    journalImpactType_list = Literal['CiteScore', 'SNIP', 'SJR']
    indexType_list = Literal['hIndex', 'h5Index', 'gIndex', 'mIndex']

    @property
    def name(self):
        return chained_get(self._results, ['author', 'name'], 'Unknown')
    
    @property
    def id(self):
        return chained_get(self._results, ['author', 'id'])
    
    @property
    def irl(self):
        return chained_get(self._results, ['author', 'irl'])

    @property
    def dataSource(self):
        return [[key, value] for key, value in self._dataSource.items()]


    def __init__(self,
                 author_id: Union[int, str],
                 refresh: Union[bool, int] = False,
                 **kwds: str
                 ) -> None:
        """Interaction with the Author Retrieval API.

        :param author_id: The ID or the EID of the author.
        :param refresh: Whether to refresh the cached file if it exists or not.
                        If int is passed, cached file will be refreshed if the
                        number of days since last modification exceeds that value.
        :param view: The view of the file that should be downloaded.  Allowed
                     values: METRICS, LIGHT, STANDARD, ENHANCED, where STANDARD
                     includes all information of LIGHT view and ENHANCED
                     includes all information of any view.  For details see
                     https://dev.elsevier.com/sc_author_retrieval_views.html.
                     Note: Neither the BASIC nor the DOCUMENTS view are active,
                     although documented.
        :param kwds: Keywords passed on as query parameters.  Must contain
                     fields and values mentioned in the API specification at
                     https://dev.elsevier.com/documentation/AuthorRetrievalAPI.wadl.

        Raises
        ------
        ValueError
            If any of the parameters `refresh` or `view` is not
            one of the allowed values.

        Notes
        -----
        The directory for cached results is `{path}/ENHANCED/{author_id}`,
        where `path` is specified in your configuration file, and `author_id`
        is stripped of an eventually leading `'9-s2.0-'`.
        """
        # Load json
        self._id = str(author_id).split('-')[-1]
        self._refresh = refresh

        Lookup.__init__(self,
                        api='AuthorLookup',
                        identifier=self._id,
                        complement="metrics",
                        **kwds)
        self.kwds = kwds

        # Parse json
        self._results = self._json['results'][0]
        self._dataSource = self._json['dataSource']


    def __str__(self) -> str:
        """Return a summary string."""
        date = self.get_cache_file_mdate().split()[0]
        s = f"Your choice, as of {date}:\n"\
            f"\t- Name: {self.name}\n"\
            f"\t- ID: \t{self.id}"
        return s
    

    def _get_metrics_rawdata(self, 
                    author_ids: str = '',
                    metricType: metricType_liste = 'ScholarlyOutput',
                    yearRange: yearRange_list = '5yrs',
                    subjectAreaFilterURI: str = '',
                    includeSelfCitations: bool = True,
                    byYear: bool = True,
                    includedDocs: includedDocs_list = 'AllPublicationTypes',
                    journalImpactType: journalImpactType_list = 'CiteScore',
                    showAsFieldWeighted: bool = False,
                    indexType: indexType_list = 'hIndex') -> any:
        
        author_ids = self._id if author_ids == '' else author_ids

        params = {
            "authors": author_ids,
            "metricTypes": metricType,
            "yearRange": yearRange,
            "subjectAreaFilterURI": subjectAreaFilterURI,
            "includeSelfCitations": includeSelfCitations,
            "byYear": byYear,
            "includedDocs": includedDocs,
            "journalImpactType": journalImpactType,
            "showAsFieldWeighted": showAsFieldWeighted,
            "indexType": indexType
        }

        response = get_content(url=URLS['AuthorLookup']+'metrics', api='AuthorLookup', params=params, **self.kwds)
        data = response.json()['results'][0]['metrics'][0]
        last_key = list(data.keys())[-1]
        return data[last_key]
    
    def get_metrics_Collaboration(self, 
                author_ids: str = '',
                metricType: Literal['AcademicCorporateCollaboration', 'AcademicCorporateCollaborationImpact', 'Collaboration', 'CollaborationImpact'] = 'AcademicCorporateCollaboration',
                collabType: Literal['Academic-corporate collaboration', 'No academic-corporate collaboration', 'Institutional collaboration', 'International collaboration', 'National collaboration', 'Single authorship'] = 'No academic-corporate collaboration',
                value_or_percentage: Literal['valueByYear', 'percentageByYear'] = 'valueByYear', 
                yearRange: yearRange_list = '5yrs',
                subjectAreaFilterURI: str = '',
                includeSelfCitations: bool = True,
                byYear: bool = True,
                includedDocs: includedDocs_list = 'AllPublicationTypes',
                journalImpactType: journalImpactType_list = 'CiteScore',
                showAsFieldWeighted: bool = False,
                indexType: indexType_list = 'hIndex'):
        if metricType in ('AcademicCorporateCollaboration', 'AcademicCorporateCollaborationImpact'):
            self._check_args(collabType, metricType, ('Academic-corporate collaboration', 'No academic-corporate collaboration'))
        elif metricType in ('Collaboration', 'CollaborationImpact'):
            self._check_args(collabType, metricType, ('Institutional collaboration', 'International collaboration', 'National collaboration', 'Single authorship'))
        if metricType in ('AcademicCorporateCollaborationImpact', 'CollaborationImpact'):
            value_or_percentage = 'valueByYear'

        return MetricsFormatage(self._for_advanced_metrics(metricType,yearRange,subjectAreaFilterURI,includeSelfCitations,byYear,includedDocs,journalImpactType,showAsFieldWeighted,indexType,
                                                    author_ids, 'collabType', collabType, value_or_percentage))
    
    def get_metrics_Percentile  (self, 
                author_ids: str = '',
                metricType: Literal['PublicationsInTopJournalPercentiles', 'OutputsInTopCitationPercentiles'] = 'OutputsInTopCitationPercentiles',
                threshold: Literal[1, 5, 10, 25] = 10,
                value_or_percentage: Literal['valueByYear', 'percentageByYear'] = 'valueByYear',
                yearRange: yearRange_list = '5yrs',
                subjectAreaFilterURI: str = '',
                includeSelfCitations: bool = True,
                byYear: bool = True,
                includedDocs: includedDocs_list = 'AllPublicationTypes',
                journalImpactType: Literal['CiteScore', 'SNIP', 'SJR'] = 'CiteScore',
                showAsFieldWeighted: bool = False,
                indexType: Literal['hIndex', 'h5Index', 'gIndex', 'mIndex'] = 'hIndex'):
        return MetricsFormatage(self._for_advanced_metrics(metricType,yearRange,subjectAreaFilterURI,includeSelfCitations,byYear,includedDocs,journalImpactType,showAsFieldWeighted,indexType,
                                                    author_ids, 'threshold', threshold, value_or_percentage))
    
    def get_metrics_Other  (self, 
                author_ids: str = '',
                metricType: Literal['CitationCount', 'CitationsPerPublication', 'CitedPublications', 'FieldWeightedCitationImpact', 'ScholarlyOutput'] = 'ScholarlyOutput',
                yearRange: yearRange_list = '5yrs',
                subjectAreaFilterURI: str = '',
                includeSelfCitations: bool = True,
                byYear: bool = True,
                includedDocs: includedDocs_list = 'AllPublicationTypes',
                journalImpactType: Literal['CiteScore', 'SNIP', 'SJR'] = 'CiteScore',
                showAsFieldWeighted: bool = False,
                indexType: Literal['hIndex', 'h5Index', 'gIndex', 'mIndex'] = 'hIndex'):
        return MetricsFormatage({'valueByYear': self._for_advanced_metrics(metricType,yearRange,subjectAreaFilterURI,includeSelfCitations,byYear,includedDocs,journalImpactType,showAsFieldWeighted,indexType,
                                                    author_ids)})


    def _get_institution_rawdata(self,
                                institutionId: Union[str, int],
                                yearRange: yearRange_list = '5yrs',
                                limit: int = 500,
                                offset: int = 0):
        
        # institutionId = self._default_current_institution_id if institutionId == 0 else str(institutionId)
        limit = 1 if limit < 1 else (500 if limit > 500 else limit)

        params = {
            "yearRange": yearRange,
            "limit": limit,
            "offset": offset
        }

        response = get_content(url=URLS['AuthorLookup']+'institutionId/'+str(institutionId), api='AuthorLookup', params=params, **self.kwds)
        data = response.json()
        del data['link']
        return data

    def institutional_authors(self, institutionId: Union[str, int], yearRange: yearRange_list = '5yrs'):
        data_list = self._get_institution_rawdata(institutionId=institutionId, yearRange=yearRange)['authors']
        result_dict = {item.get('id'): item for item in data_list if item.get('id') is not None}
        for item in result_dict.values():
            item.pop('link', None)
            item.pop('uri', None)
            item.pop('id', None)
        return InstitutionalFormatage(result_dict)
    
    def institutional_total_count(self, institutionId: Union[str, int], yearRange: yearRange_list = '5yrs'):
        return self._get_institution_rawdata(institutionId=institutionId, yearRange=yearRange)['totalCount']



    def _check_args(self, collabType: tuple, metricType: str, valid_collab_types):
        if collabType not in valid_collab_types:
            raise ValueError(f"Invalid collabType '{collabType}' for metricType '{metricType}'. "
                             f"Valid collabTypes are: {', '.join(valid_collab_types)}")

    def _for_advanced_metrics(self, metricType, yearRange, subjectAreaFilterURI, includeSelfCitations, byYear, includedDocs, journalImpactType, showAsFieldWeighted, indexType,
                                author_ids: str, key_name: str = '', value: Union[int, str] = '', value_or_percentage: str = 'valueByYear'):
        author_ids = self._id if author_ids == '' else author_ids
        data_dict = self._get_metrics_rawdata(author_ids, metricType, yearRange, subjectAreaFilterURI, includeSelfCitations, byYear, includedDocs, journalImpactType, showAsFieldWeighted, indexType)
        
        result_dict = {}
        if not key_name == '':
            for data in data_dict:
                collab_type = data[key_name]
                if collab_type == value:
                    result_dict[key_name] = collab_type
                    result_dict[value_or_percentage] = data[value_or_percentage]
        else:
            result_dict = data_dict

        return result_dict



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
        return str("You must select a property: 'Raw', 'List', 'Dictionary', 'DataFrame' after calling the AuthorLookup class method!")    
    

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
        return str("You must select a property: 'List', 'Dictionary', 'DataFrame' after calling the AuthorLookup class method!")
    
