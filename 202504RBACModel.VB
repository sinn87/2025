' VB.NET RBAC 模型示例（支持角色继承、权限约束与条件访问控制）

Module RBACModel

    ' 定义 Permission 类（表示权限）
    Public Class Permission
        Public Property Name As String

        Public Sub New(name As String)
            Me.Name = name
        End Sub
    End Class

    ' 定义 Role 类（表示角色，支持角色继承）
    Public Class Role
        Public Property Name As String
        Public Property Permissions As New List(Of Permission)
        Public Property ParentRoles As New List(Of Role) ' 父角色列表
        Public Property ConflictRoles As New List(Of Role) ' 冲突角色列表

        Public Sub New(name As String)
            Me.Name = name
        End Sub

        Public Sub AddPermission(permission As Permission)
            Permissions.Add(permission)
        End Sub

        Public Sub AddParentRole(parentRole As Role)
            ParentRoles.Add(parentRole)
        End Sub

        Public Sub AddConflictRole(conflictRole As Role)
            ConflictRoles.Add(conflictRole)
        End Sub

        ' 获取所有权限（包括继承的权限）
        Public Function GetAllPermissions() As List(Of Permission)
            Dim allPermissions As New List(Of Permission)(Permissions)

            For Each parentRole As Role In ParentRoles
                allPermissions.AddRange(parentRole.GetAllPermissions())
            Next

            Return allPermissions
        End Function
    End Class

    ' 定义 User 类（表示用户）
    Public Class User
        Public Property Name As String
        Public Property Roles As New List(Of Role)

        Public Sub New(name As String)
            Me.Name = name
        End Sub

        Public Sub AddRole(role As Role)
            ' 检查冲突角色
            For Each existingRole In Roles
                If existingRole.ConflictRoles.Contains(role) OrElse role.ConflictRoles.Contains(existingRole) Then
                    Console.WriteLine($"角色冲突：用户 {Name} 不能同时拥有 {existingRole.Name} 和 {role.Name}。")
                    Return
                End If
            Next
            Roles.Add(role)
        End Sub

        ' 检查用户是否拥有某个权限（增加条件访问控制）
        Public Function HasPermission(permissionName As String, Optional timeOfDay As String = "") As Boolean
            For Each role As Role In Roles
                For Each permission As Permission In role.GetAllPermissions()
                    If permission.Name = permissionName Then
                        ' 条件访问控制：只能在白天访问
                        If timeOfDay = "night" AndAlso permissionName = "编辑报告" Then
                            Console.WriteLine($"权限拒绝：{Name} 无法在夜晚进行编辑操作。")
                            Return False
                        End If
                        Return True
                    End If
                Next
            Next
            Return False
        End Function
    End Class

    Sub Main()
        ' 创建权限
        Dim viewReport As New Permission("查看报告")
        Dim editReport As New Permission("编辑报告")
        Dim deleteReport As New Permission("删除报告")

        ' 创建角色
        Dim guest As New Role("访客")
        guest.AddPermission(viewReport)

        Dim editor As New Role("编辑者")
        editor.AddPermission(editReport)
        editor.AddParentRole(guest)

        Dim admin As New Role("管理员")
        admin.AddPermission(deleteReport)
        admin.AddParentRole(editor)

        ' 创建冲突角色
        Dim auditor As New Role("审计员")
        auditor.AddPermission(viewReport)
        editor.AddConflictRole(auditor)

        ' 创建用户
        Dim user1 As New User("小李")
        user1.AddRole(guest)

        Dim user2 As New User("小王")
        user2.AddRole(editor)
        user2.AddRole(auditor)

        Dim user3 As New User("小张")
        user3.AddRole(admin)

        ' 测试用户权限（加入条件访问控制）
        Console.WriteLine($"{user1.Name} 是否可以查看报告？{user1.HasPermission("查看报告")}")
        Console.WriteLine($"{user1.Name} 是否可以编辑报告？{user1.HasPermission("编辑报告")}")

        Console.WriteLine($"{user2.Name} 是否可以查看报告？{user2.HasPermission("查看报告")}")
        Console.WriteLine($"{user2.Name} 是否可以编辑报告？{user2.HasPermission("编辑报告", "night")}") ' 测试条件访问控制
        Console.WriteLine($"{user2.Name} 是否可以删除报告？{user2.HasPermission("删除报告")}")

        Console.WriteLine($"{user3.Name} 是否可以删除报告？{user3.HasPermission("删除报告")}")
        Console.WriteLine($"{user3.Name} 是否可以编辑报告？{user3.HasPermission("编辑报告", "day")}")

        Console.ReadLine()
    End Sub

End Module
