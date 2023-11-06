# smartsheet-helper
Simple class that helps dealing with smarsheet api, like dealing with columin ids and types

it uses environment variables for keys and ids:

SMARTSHEET_SHEET_ID=your_sheet_id
SMARTSHEET_ACCESS_TOKEN=your_secret_access_token

## Danjgo webhook handler example

'''
from django.views.decorators.http import require_http_methods
from django.views.decorators.csrf import csrf_exempt
from django.http import JsonResponse

@require_http_methods(["POST"])
@csrf_exempt
def handleWebhook(request):
    #handle challenge
    if("Smartsheet-Hook-Challenge" in request.headers):
        res = {"smartsheetHookResponse": request.headers["Smartsheet-Hook-Challenge"]}
        return JsonResponse(res)
    data = json.loads(request.body)
    # data = {
    #     "nonce":"6939c036-b01b-1111-2222-4e085ef1a5a0",
    #     "timestamp":"2023-11-05T13:07:26.115+00:00",
    #     "webhookId":6847500251111108,
    #     "scope":"sheet",
    #     "scopeObjectId":6817504622273380,
    #     "events":[
    #         {
    #             "objectType":"cell",
    #             "eventType":"updated",
    #             "rowId":5849931133364004,
    #             "columnId":3996433355807364,
    #             "userId":8583582333729092,
    #             "timestamp":"2023-11-05T13:07:20.000+00:00"
    #         }
    #     ]
    # }

    # do something

    return JsonResponse({"result": "success"})

'''