from os import environ
from pathlib import Path

# Paths for cached files
if (Path.home()/".scopus").exists():
    BASE_PATH_SCOPUS = Path.home()/".scopus"
elif (Path.home()/".pybliometrics"/"Scopus").exists():
    BASE_PATH_SCOPUS = Path.home()/".pybliometrics"/"Scopus"
else:
    BASE_PATH_SCOPUS = Path.home()/".cache"/"pybliometrics"/"Scopus"

if (Path.home()/".scival").exists():
    BASE_PATH_SCIVAL = Path.home()/".scival"
elif (Path.home()/".pybliometrics"/"SciVal").exists():
    BASE_PATH_SCIVAL = Path.home()/".pybliometrics"/"SciVal"
else:
    BASE_PATH_SCIVAL = Path.home()/".cache"/"pybliometrics"/"SciVal"

DEFAULT_PATHS = {
    'AbstractRetrieval': BASE_PATH_SCOPUS/'abstract_retrieval',
    'AffiliationRetrieval': BASE_PATH_SCOPUS/'affiliation_retrieval',
    'AffiliationSearch': BASE_PATH_SCOPUS/'affiliation_search',
    'AuthorRetrieval': BASE_PATH_SCOPUS/'author_retrieval',
    'AuthorSearch': BASE_PATH_SCOPUS/'author_search',
    'CitationOverview': BASE_PATH_SCOPUS/'citation_overview',
    'ScopusSearch': BASE_PATH_SCOPUS/'scopus_search',
    'SerialSearch': BASE_PATH_SCOPUS/'serial_search',
    'SerialTitle': BASE_PATH_SCOPUS/'serial_title',
    'PlumXMetrics': BASE_PATH_SCOPUS/'plumx',
    'SubjectClassifications': BASE_PATH_SCOPUS/'subject_classification',

    'AuthorLookup': BASE_PATH_SCIVAL/'author_lookup',
    'CountryLookup': BASE_PATH_SCIVAL/'country_lookup',
    'CountryGroupLookup': BASE_PATH_SCIVAL/'country_group_lookup',
    'InstitutionLookup': BASE_PATH_SCIVAL/'institution_lookup',
    'InstitutionGroupLookup': BASE_PATH_SCIVAL/'author_group_lookup',
    'PublicationLookup': BASE_PATH_SCIVAL/'publication_lookup',
    'ScopusSourceLookup': BASE_PATH_SCIVAL/'scopus_source_lookup',
    'SubjectAreaLookup': BASE_PATH_SCIVAL/'subject_area_lookup',
    'TopicLookup': BASE_PATH_SCIVAL/'topic_lookup',
    'TopicClusterLookup': BASE_PATH_SCIVAL/'topic_cluster_lookup',
    'WorldLookup': BASE_PATH_SCIVAL/'world_lookup'
}

# Configuration file location
if 'PYB_CONFIG_FILE' in environ:
    CONFIG_FILE = Path(environ['PYB_CONFIG_FILE'])
else:
    if (Path.home()/".scopus").exists():
        CONFIG_FILE = Path.home()/".scopus"/"config.ini"
    elif (Path.home()/".pybliometrics"/"config.ini").exists():
        CONFIG_FILE = Path.home()/".pybliometrics"/"config.ini"
    else:
        CONFIG_FILE = Path.home()/".config"/"pybliometrics.cfg"

# URLs for all classes
RETRIEVAL_BASE = 'https://api.elsevier.com/content/'
SEARCH_BASE = 'https://api.elsevier.com/content/search/'
LOOKUP_BASE = 'https://api.elsevier.com/analytics/scival/'
URLS = {
    'AbstractRetrieval': RETRIEVAL_BASE + 'abstract/',
    'AffiliationRetrieval': RETRIEVAL_BASE + 'affiliation/affiliation_id/',
    'AffiliationSearch': SEARCH_BASE + 'affiliation',
    'AuthorRetrieval': RETRIEVAL_BASE + 'author/author_id/',
    'AuthorSearch': SEARCH_BASE + 'author',
    'CitationOverview': RETRIEVAL_BASE + 'abstract/citations/',
    'ScopusSearch': SEARCH_BASE + 'scopus',
    'SerialSearch': RETRIEVAL_BASE + 'serial/title',
    'SerialTitle': RETRIEVAL_BASE + 'serial/title/issn/',
    'SubjectClassifications': RETRIEVAL_BASE + 'subject/scopus',
    'PlumXMetrics': 'https://api.elsevier.com/analytics/plumx/',

    'AuthorLookup': LOOKUP_BASE + 'author/',
    'CountryLookup': LOOKUP_BASE + 'country/',
    'CountryGroupLookup': LOOKUP_BASE + 'countryGroup/',
    'InstitutionLookup': LOOKUP_BASE + 'institution/',
    'InstitutionGroupLookup': LOOKUP_BASE + 'institutionGroup/',
    'PublicationLookup': LOOKUP_BASE + 'publication/',
    'ScopusSourceLookup': LOOKUP_BASE + 'scopusSource/',
    'SubjectAreaLookup': LOOKUP_BASE + 'subjectArea/',
    'TopicLookup': LOOKUP_BASE + 'topic/',
    'TopicClusterLookup': LOOKUP_BASE + 'topicCluster/',
    'WorldLookup': LOOKUP_BASE + 'world/'
}

# Throttling limits (in queries per second) // 0 = no limit
RATELIMITS = {
    'AbstractRetrieval': 9,
    'AffiliationRetrieval': 9,
    'AffiliationSearch': 6,
    'AuthorRetrieval': 3,
    'AuthorSearch': 2,
    'CitationOverview': 4,
    'ScopusSearch': 9,
    'SerialSearch': 6,
    'SerialTitle': 6,
    'PlumXMetrics': 6,
    'SubjectClassifications': 0,

    'AuthorLookup': 6,
    'CountryLookup': 6,
    'CountryGroupLookup': 6,
    'InstitutionLookup': 6,
    'InstitutionGroupLookup': 6,
    'PublicationLookup': 6,
    'ScopusSourceLookup': 6,
    'SubjectAreaLookup': 6,
    'TopicLookup': 6,
    'TopicClusterLookup': 6,
    'WorldLookup': 6
}

# Other API restrictions
SEARCH_MAX_ENTRIES = 5_000
