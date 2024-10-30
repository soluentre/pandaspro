import pandas as pd
import numpy as np

# 设置随机种子以确保可重复性
np.random.seed(42)

# 样本数据集的行数
num_rows = 500

# 创建员工姓名
first_names = ["John", "Jane", "Sam", "Susan", "Michael", "Mary", "James", "Linda", "Robert", "Patricia"]
last_names = ["Smith", "Johnson", "Brown", "Taylor", "Anderson", "Thomas", "Jackson", "White", "Harris", "Martin"]
names = [f"{fn} {ln}" for fn in first_names for ln in last_names]
employee_names = np.random.choice(names, num_rows, replace=True)

# 创建员工编号（唯一）
employee_ids = np.arange(1000, 1000 + num_rows)

# 创建部门
departments = ["HR", "Finance", "IT", "Operations"]
employee_departments = np.random.choice(departments, num_rows, replace=True)

# 创建子部门
sub_departments = {
    "HR": ["Recruitment", "Employee Relations", "Training"],
    "Finance": ["Accounting", "Controlling", "Auditing"],
    "IT": ["Development", "Support", "Security"],
    "Operations": ["Logistics", "Manufacturing", "Procurement"],
}
employee_sub_departments = [np.random.choice(sub_departments[dept]) for dept in employee_departments]

# 创建员工级别
levels = ["Junior", "Senior", "Lead"]
employee_levels = np.random.choice(levels, num_rows, replace=True)

# 创建性别
genders = ["Male", "Female"]
employee_genders = np.random.choice(genders, num_rows, replace=True)

# 创建年龄
employee_ages = np.random.randint(22, 65, num_rows)

# 创建工作地
locations = ["New York", "San Francisco", "Chicago", "Seattle"]
employee_locations = np.random.choice(locations, num_rows, replace=True)

# 创建薪酬
salary_ranges = {
    "Junior": (40000, 60000),
    "Mid": (60000, 80000),
    "Senior": (80000, 120000),
    "Lead": (120000, 180000)
}
employee_salaries = [np.random.randint(salary_ranges[level][0], salary_ranges[level][1]) for level in employee_levels]

# 创建数据框
df = pd.DataFrame({
    "id": employee_ids,
    "name": employee_names,
    "department": employee_departments,
    "unit": employee_sub_departments,
    "grade": employee_levels,
    "gender": employee_genders,
    "age": employee_ages,
    "location": employee_locations,
    "salary": employee_salaries
})

