function SearchAzGraph {
    [CmdletBinding()]
    param(
        [string]
        $Query
    )

    try {
        $results = [System.Collections.ArrayList]::new()

        $responses = Search-AzGraph -Query $query -First 100

        $results.AddRange([array]$responses)

        while ($responses.SkipToken) {
            $responses = Search-AzGraph -Query $query -SkipToken $responses.SkipToken

            $results.AddRange([array]$responses)
        }

        return $results.ToArray()
    } catch {
        throw $PSItem
    }
}
