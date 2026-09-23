SELECT p.title as title,p.desired_result as desired_result , a.todo as todo
FROM action_current As a
    LEFT JOIN project_current As p
        ON p.id=a.project_id
WHERE a.done=False
ORDER BY p.id,a.id