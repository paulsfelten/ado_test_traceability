from datetime import datetime, timedelta
import requests
import json
import csv
import argparse

parse = argparse.ArgumentParser(description="ForTest!")
parse.add_argument('-token', '--token')
parse.add_argument('-project', '--project')
parse.add_argument('-organization', '--organization')
args = parse.parse_args()

def get_stories_bugs_in_iteration(organization, project, personal_access_token):
    url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/wiql?api-version=6.0"
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Basic {personal_access_token}'
    }
    wiql_query = {
        "query": f"""
        SELECT
            [System.Id],
            [System.WorkItemType],
            [System.Title],
            [System.State],
            [System.Tags],
            [System.AreaPath]
        FROM workitems
        WHERE
            [System.TeamProject] = '{project}'
            AND (
                ([System.WorkItemType] = 'User Story'
                OR [System.WorkItemType] = 'Bug')
                AND [System.State] <> ''
                AND (
                    [System.IterationPath] = @currentIteration('todo')
                    OR [System.IterationPath] = @currentIteration('todo')
                    OR [System.IterationPath] = @currentIteration('todo')
                    OR [System.IterationPath] = @currentIteration('todo')
                    OR [System.IterationPath] = @currentIteration('todo')-1
                    OR [System.IterationPath] = @currentIteration('todo')-1
                    OR [System.IterationPath] = @currentIteration('todo')-1
                    OR [System.IterationPath] = @currentIteration('todo')-1
                )
                AND (
                    [System.AreaPath] = '{project}\todo'
                    OR [System.AreaPath] = '{project}\todo'
                    OR [System.AreaPath] = '{project}\todo'
                    OR [System.AreaPath] = '{project}\todo'
                )
            )
        ORDER BY [System.AreaPath]
        """
    }
    response = requests.post(url, headers=headers, data=json.dumps(wiql_query))
    if response.status_code == 200:
        work_items = response.json().get('workItems', [])
        return work_items
    else:
        print(f"Failed to retrieve work items. Status code: {response.status_code}")
        print(f"Response: {response.text}")
        return None

def get_all_test_plan_tests(organization, project, personal_access_token):
    url = f'https://dev.azure.com/{organization}/{project}/_apis/test/plans?api-version=5.0'

    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Basic {personal_access_token}'
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        test_plans = response.json()['value']
        current_year = datetime.now().year
        test_plans_with_test_cases = [] # <test plan>:<test case>

        for plan in test_plans:
            if plan['name'].startswith(str(current_year)):
                date_part = plan['name'].split('_')[0]
                try:
                    parsed_date = datetime.strptime(date_part, '%Y%m%d')

                    if parsed_date.date() < (datetime.now().date() + timedelta(days=30)) and parsed_date.date() > (datetime.now().date() - timedelta(days=30)):


                        plan_id = plan['id']
                        plan_url = f'https://dev.azure.com/{organization}/{project}/_apis/test/Plans/{plan_id}/Suites?api-version=5.0'
                        plan_response = requests.get(plan_url, headers=headers)
                        
                        if plan_response.status_code == 200:
                            test_suites = plan_response.json()['value']
                            for suite in test_suites:
                                suite_id = suite['id']
                                suite_url = f'https://dev.azure.com/{organization}/{project}/_apis/test/Plans/{plan_id}/Suites/{suite_id}/points?api-version=7.0'
                                suite_response = requests.get(suite_url, headers=headers)
                                if suite_response.status_code == 200:
                                    test_cases = suite_response.json()['value']
                                    for test_case in test_cases:
                                        test_plans_with_test_cases.append(f'{plan["name"]} ({test_case["outcome"]}):{test_case["testCase"]["id"]}')
                except:
                    continue
        return test_plans_with_test_cases
    else:
        return None

def get_linked_test_cases(organization, project, work_item_id, personal_access_token):
    url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{work_item_id}?$expand=1&api-version=7.0"
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Basic {personal_access_token}'
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        fields = response.json().get('fields', {})
        return_fields = [ fields['System.AreaPath'], fields['System.Title'], fields['System.State'], fields['System.IterationPath'] , fields['System.WorkItemType'] ]

        relations = response.json().get('relations', [])

        work_item_ids = []
        for rel in relations:
            if rel['rel'] == 'System.LinkTypes.Hierarchy-Forward' or rel['rel'] == 'System.LinkTypes.Hierarchy-Reverse' or rel['rel'] == 'Microsoft.VSTS.Common.TestedBy-Reverse' or rel['rel'] == 'Microsoft.VSTS.Common.TestedBy-Forward':
                work_item_ids.append(rel['url'].split('/')[-1])
        work_item_ids_str = ', '.join(map(str, work_item_ids))

        if work_item_ids:
            wiql_query = {
                "query": f"""
                SELECT [System.Id]
                FROM workitems
                WHERE
                    [System.TeamProject] = '{project}'
                    AND [System.WorkItemType] = 'Test Case'
                    AND [System.Id] IN ({work_item_ids_str})
                """
            }

            url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/wiql?api-version=7.0"
            headers = {
                'Content-Type': 'application/json',
                'Authorization': f'Basic {personal_access_token}'
            }
            response = requests.post(url, json=wiql_query, headers=headers)
            if response.status_code == 200:
                work_items = response.json().get('workItems', [])
                test_cases = [item['id'] for item in work_items if item['id']]
                return test_cases, return_fields
            else:
                return None, return_fields
        else:
            return None, return_fields
    else:
        return None, None

def get_test_plans(test_case_ids):
    plans = []
    if test_case_ids:
        for test_case_id in test_case_ids:
            for item in all_test_cases:
                if str(test_case_id) in item:
                    plans.append(item.split(':')[0])
        return ', '.join(map(str, plans)) 
    else:
        return 'None'

organization = args.organization
project = args.project
encoded_pat = args.token

print("Getting user stories...")
stories = get_stories_bugs_in_iteration(organization, project, encoded_pat)
print("Getting all test plans and test cases...")
all_test_cases = get_all_test_plan_tests(organization, project, encoded_pat)

# Write results to a CSV file
with open('qa_traceability.csv', mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['Story/Bug ID', 'Type', 'Area Path', 'Title', 'State', 'Iteration', 'Linked Test Cases', 'Test Plan'])

    if stories:
        print("Finding linked test cases and user story meta data...")
        for user_story in stories:
            linked_test_cases, return_fields = get_linked_test_cases(organization, project, user_story['id'], encoded_pat)
            linked_test_cases_str = ', '.join(map(str, linked_test_cases)) if linked_test_cases else 'None'

            test_plans_str = get_test_plans(linked_test_cases)
                
            writer.writerow([user_story['id'], return_fields[4], return_fields[0], return_fields[1], return_fields[2], return_fields[3], linked_test_cases_str, test_plans_str])
    else:
        print("No user stories or bugs found.")
print("Done!")
