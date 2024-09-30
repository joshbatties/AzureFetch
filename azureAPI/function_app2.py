import logging
import azure.functions as func
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import os
import json

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    # Authenticate the request
    if not req.headers.get('Authorization'):
        return func.HttpResponse(
            'Unauthorized',
            status_code=401
        )

    # Extract parameters from the request
    container_number = req.params.get('containerNumber')
    company_code = req.params.get('companyCode')

    if not container_number and not company_code:
        return func.HttpResponse(
            "Please pass a containerNumber or companyCode on the query string",
            status_code=400
        )

    # SharePoint credentials and site URL (using Client ID and Secret)
    site_url = os.environ['SP_SITE_URL']
    client_id = os.environ['SP_CLIENT_ID']
    client_secret = os.environ['SP_CLIENT_SECRET']

    # Authenticate with SharePoint using Client ID and Secret
    ctx_auth = ClientCredential(client_id, client_secret)
    ctx = ClientContext(site_url).with_credentials(ctx_auth)

    # Access the SharePoint list
    list_object = ctx.web.lists.get_by_title('ContainerData')

    # Build the query based on containerNumber or companyCode
    from office365.sharepoint.listitems.caml_query import CamlQuery
    caml_query = CamlQuery()
    if container_number:
        caml_query.ViewXml = f"""
            <View>
                <Query>
                    <Where>
                        <Eq>
                            <FieldRef Name='ContainerNumber'/>
                            <Value Type='Text'>{container_number}</Value>
                        </Eq>
                    </Where>
                </Query>
            </View>
        """
    elif company_code:
        caml_query.ViewXml = f"""
            <View>
                <Query>
                    <Where>
                        <Eq>
                            <FieldRef Name='CompanyCode'/>
                            <Value Type='Text'>{company_code}</Value>
                        </Eq>
                    </Where>
                </Query>
            </View>
        """

    # Execute the query and fetch items
    items = list_object.get_items(caml_query)
    ctx.load(items)
    ctx.execute_query()

    # Prepare the response with the retrieved data
    data = [item.properties for item in items]

    return func.HttpResponse(
        json.dumps(data),
        status_code=200,
        mimetype="application/json"
    )

