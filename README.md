Tool use HOWTO

Prereqs: Excel files are required to have the following properties
- Row 1 has Column title data. Each column must hace a unique title 

1. Install a semi-recent version of the Python programming language (version 3.7 or later should work, previous versions *may* work)

2. Install the following additional Python modules if you do not already have them:
  pip3 install toml (this one comes build-in in versions 3.11 or later)
  pip3 install openpyxl

3. Edit the gbiftaxonfetch.toml/gbiftaxonfetch.toml setup file. You will probably need to set at least the following properties:
- taxonomy.class. The GBIF taxon search is limited to a single class (for now) to avoid problems with duplicate names between animals/plants.
- excel.col_taxon_name
  excel.col_taxon_author
  excel.col_taxon_rank
  there are the titles of columns containing the taxon name, taxon authorship, and taxon rank (species,genus...) in the input file
- gbif.cachefile.
  To speed up the process and take load off GBIF servers, results are stored locally for future reuse. This should point to a writeable location on your own computer.
  If local caching is not wanted, set gbif.flush_GBIF_cache = true.
  To reset local data, delete the cache file or set gbif.flush_GBIF_cache = true for one run.
