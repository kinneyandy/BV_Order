# Import the contents of the OrderTest.csv file and store it in the $order_list variable.
$order_list = Import-Csv .\Orders.csv
# This is where the document will be saved:
$xmlFile = "C:\Users\akinney\Documents\Powershell\bv_ppe_tag_feed_VaxUK_$((Get-Date).ToString("yyyyMMdd_HHmmss")).xml"

$xmlWriter = New-Object System.XMl.XmlTextWriter($xmlFile,$Encoding.UTF8)

# choose a pretty formatting:
$xmlWriter.Formatting = 'Indented'
$xmlWriter.Indentation = 1
$XmlWriter.IndentChar = "`t"


$xmlWriter.WriteStartDocument()
# Start Element - Feed
$xmlWriter.WriteStartElement("Feed");
$xmlWriter.WriteAttributeString("xmlns", "http://www.bazaarvoice.com/xs/PRR/PostPurchaseFeed/14.7")
# Variable used to identify different order by comparing to previous value
$PrevUserID = ''
# Loop through all the records in the CSV
foreach ($orderline in $order_list) {
    
    # Check if the current record's UserID is equal to the value of the UserID parameter.
    if ($orderline.UserID -eq $PrevUserID) {

        # Start Element - Product
        $xmlWriter.WriteStartElement("Product");
        $xmlWriter.WriteElementString('ExternalId',$orderline.ExternalId)
        $xmlWriter.WriteElementString('Name',$orderline.Name)
        $xmlWriter.WriteElementString('Price',$orderline.Price)
        # End Element - Product
        $xmlWriter.WriteEndElement()

    }elseif($PrevUserID -eq ''){

        # Start Element - Interaction
        $xmlWriter.WriteStartElement("Interaction")
        $xmlWriter.WriteElementString('TransactionDate',$orderline.TransactionDate)
        $xmlWriter.WriteElementString('EmailAddress',$orderline.EmailAddress)
        $xmlWriter.WriteElementString('UserName',$orderline.UserName)
        $xmlWriter.WriteElementString('UserID',$orderline.UserID)
        # Start Element - Products
        $xmlWriter.WriteStartElement("Products")
        # Start Element - Product
        $xmlWriter.WriteStartElement("Product")
        $xmlWriter.WriteElementString('ExternalId',$orderline.ExternalId)
        $xmlWriter.WriteElementString('Name',$orderline.Name)
        $xmlWriter.WriteElementString('Price',$orderline.Price)
        # End Element - Product
        $xmlWriter.WriteEndElement()

    }else {
        # End Element - Products
        $xmlWriter.WriteEndElement()
        # End Element - Interaction
        $xmlWriter.WriteEndElement()
        # Start Element - Interaction
        $xmlWriter.WriteStartElement("Interaction");
        $xmlWriter.WriteElementString('TransactionDate',$orderline.TransactionDate)
        $xmlWriter.WriteElementString('EmailAddress',$orderline.EmailAddress)
        $xmlWriter.WriteElementString('UserName',$orderline.UserName)
        $xmlWriter.WriteElementString('UserID',$orderline.UserID)
        # Start Element - Products
        $xmlWriter.WriteStartElement("Products");
        # Start Element - Product
        $xmlWriter.WriteStartElement("Product");
        $xmlWriter.WriteElementString('ExternalId',$orderline.ExternalId)
        $xmlWriter.WriteElementString('Name',$orderline.Name)
        $xmlWriter.WriteElementString('Price',$orderline.Price)
        # End Element - Product
        $xmlWriter.WriteEndElement()

    }

    $PrevUserID = $orderline.UserID

}
# End Element - Products
$xmlWriter.WriteEndElement()
# End Element - Interaction
$xmlWriter.WriteEndElement()
# End Element - Feed
$xmlWriter.WriteEndElement()

# finalize the document:
$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()