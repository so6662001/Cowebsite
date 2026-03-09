<template>
  <div>
    <!-- Hero -->
    <section class="gradient-hero pt-32 pb-20">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h1 class="text-4xl lg:text-5xl font-bold text-white mb-6">行业解决方案</h1>
        <p class="text-xl text-white/80 max-w-3xl mx-auto">
          针对钢贸企业和钢铁生产制造企业，提供从业务增长到降本增效的全方位数智化解决方案
        </p>
      </div>
    </section>

    <!-- Breadcrumb -->
    <nav class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4" aria-label="面包屑导航">
      <ol class="flex items-center gap-2 text-sm text-gray-500" itemscope itemtype="https://schema.org/BreadcrumbList">
        <li itemprop="itemListElement" itemscope itemtype="https://schema.org/ListItem">
          <router-link to="/" itemprop="item"><span itemprop="name">首页</span></router-link>
          <meta itemprop="position" content="1" />
        </li>
        <li class="text-gray-300">/</li>
        <li itemprop="itemListElement" itemscope itemtype="https://schema.org/ListItem">
          <span itemprop="name" class="text-gray-900 font-medium">解决方案</span>
          <meta itemprop="position" content="2" />
        </li>
      </ol>
    </nav>

    <!-- Trade Solutions -->
    <section class="py-16 lg:py-24 bg-white" aria-labelledby="trade-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="text-center mb-16">
          <span class="inline-flex items-center gap-2 px-4 py-2 bg-blue-100 text-blue-700 rounded-full text-sm font-medium mb-4">
            钢贸企业 + 生产制造企业
          </span>
          <h2 id="trade-title" class="text-3xl font-bold text-gray-900 mb-4">钢贸企业数智化解决方案</h2>
          <p class="text-lg text-steel-light max-w-2xl mx-auto">
            覆盖业务增长、降本增效、企业管理等全方位经营场景，助力钢贸企业数字化转型
          </p>
        </div>
        <div class="grid sm:grid-cols-2 lg:grid-cols-3 gap-8">
          <article v-for="solution in tradeSolutions" :key="solution.title"
            class="bg-white rounded-2xl p-8 shadow-sm border border-gray-100 card-hover">
            <div class="w-14 h-14 rounded-xl bg-blue-100 flex items-center justify-center mb-6">
              <span class="text-2xl">{{ solution.icon }}</span>
            </div>
            <h3 class="text-lg font-bold text-gray-900 mb-3">{{ solution.title }}</h3>
            <p class="text-steel-light text-sm leading-relaxed mb-4">{{ solution.desc }}</p>
            <ul class="space-y-2">
              <li v-for="point in solution.points" :key="point" class="flex items-start gap-2 text-sm text-gray-600">
                <svg class="w-4 h-4 text-primary flex-shrink-0 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                  <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/>
                </svg>
                {{ point }}
              </li>
            </ul>
          </article>
        </div>
      </div>
    </section>

    <!-- Manufacturing Solutions -->
    <section class="py-16 lg:py-24 bg-bg-section" aria-labelledby="mfg-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="text-center mb-16">
          <span class="inline-flex items-center gap-2 px-4 py-2 bg-green-100 text-green-700 rounded-full text-sm font-medium mb-4">
            钢铁生产制造企业
          </span>
          <h2 id="mfg-title" class="text-3xl font-bold text-gray-900 mb-4">钢铁生产制造数智化解决方案</h2>
          <p class="text-lg text-steel-light max-w-2xl mx-auto">
            从生产排产到质量管控，为钢铁生产制造企业提供全链条数字化管理方案
          </p>
        </div>
        <div class="grid sm:grid-cols-2 lg:grid-cols-3 gap-8">
          <article v-for="solution in mfgSolutions" :key="solution.title"
            class="bg-white rounded-2xl p-8 shadow-sm border border-gray-100 card-hover">
            <div class="w-14 h-14 rounded-xl bg-green-100 flex items-center justify-center mb-6">
              <span class="text-2xl">{{ solution.icon }}</span>
            </div>
            <h3 class="text-lg font-bold text-gray-900 mb-3">{{ solution.title }}</h3>
            <p class="text-steel-light text-sm leading-relaxed mb-4">{{ solution.desc }}</p>
            <ul class="space-y-2">
              <li v-for="point in solution.points" :key="point" class="flex items-start gap-2 text-sm text-gray-600">
                <svg class="w-4 h-4 text-green-600 flex-shrink-0 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                  <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/>
                </svg>
                {{ point }}
              </li>
            </ul>
          </article>
        </div>
      </div>
    </section>

    <!-- CTA -->
    <section class="py-16 gradient-hero text-center">
      <div class="max-w-4xl mx-auto px-4">
        <h2 class="text-3xl font-bold text-white mb-6">找到适合您的解决方案</h2>
        <p class="text-white/80 mb-8">我们的专业顾问将根据您的企业实际情况，为您推荐最合适的数智化方案</p>
        <div class="flex flex-wrap justify-center gap-4">
          <router-link to="/trial" class="px-8 py-4 bg-white text-primary font-semibold rounded-xl hover:bg-gray-100 transition-all shadow-lg">
            申请免费试用
          </router-link>
          <router-link to="/contact" class="px-8 py-4 border-2 border-white text-white font-semibold rounded-xl hover:bg-white/10 transition-all">
            咨询方案详情
          </router-link>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup>
const tradeSolutions = [
  {
    title: '业务增长解决方案',
    icon: '📈',
    desc: '通过数字化手段拓展业务渠道，整合线上线下资源，帮助钢贸企业实现业务持续增长。',
    points: ['线上获客渠道搭建', '客户资源数字化管理', '销售漏斗分析', '业务数据可视化']
  },
  {
    title: '降本增效解决方案',
    icon: '⚡',
    desc: '优化业务流程，减少人工操作，降低运营成本，提升整体经营效率。',
    points: ['业务流程自动化', '减少重复性工作', '运营成本分析', '效率提升数据追踪']
  },
  {
    title: '企业扩张解决方案',
    icon: '🚀',
    desc: '支持多分公司、多仓库、多业务线的集团化管理，助力企业快速扩张发展。',
    points: ['多分公司统一管理', '集团数据汇总分析', '权限分级管控', '标准化流程复制']
  },
  {
    title: '财务自动化解决方案',
    icon: '💰',
    desc: '实现财务核算、对账、结算自动化，大幅降低财务工作量和出错率。',
    points: ['自动对账结算', '应收应付自动管理', '智能财务报表', '多维度成本分析']
  },
  {
    title: '集团分公司管理方案',
    icon: '🏢',
    desc: '建立统一的集团管理平台，实现总部与分公司的信息互通和高效协同。',
    points: ['统一数据标准', '分级授权管理', '实时数据同步', '集团报表汇总']
  },
  {
    title: '客户/销售外出管理方案',
    icon: '👥',
    desc: '实现销售团队的外出管理、客户拜访记录、业务跟进的全流程数字化管理。',
    points: ['外出签到定位', '拜访记录管理', '客户跟进提醒', '销售业绩分析']
  },
  {
    title: '运营自动化解决方案',
    icon: '⚙️',
    desc: '实现钢贸服务商日常运营流程的自动化管理，减少人工干预，提升运营效率。',
    points: ['订单自动处理', '库存自动预警', '物流自动调度', '报表自动生成']
  },
  {
    title: '仓库管理解决方案',
    icon: '📦',
    desc: '全流程仓库数字化管理，实时掌握库存动态，提升仓库作业效率和准确率。',
    points: ['实时库存管理', '出入库自动化', '盘点效率提升', '货位优化管理']
  },
  {
    title: 'AI数字员工解决方案',
    icon: '🤖',
    desc: '利用AI技术打造企业数字员工，智能处理客服、数据分析等重复性工作，释放人力价值。',
    points: ['智能客服应答', 'AI数据分析', '自动化报告生成', '智能决策辅助']
  }
]

const mfgSolutions = [
  {
    title: '生产MES解决方案',
    icon: '🏭',
    desc: '制造执行系统全面管控生产过程，实现生产过程的可视化、数字化、智能化管理。',
    points: ['生产过程管控', '工单管理', '产出数据采集', '生产可视化']
  },
  {
    title: '生产成本管理方案',
    icon: '📉',
    desc: '精细化管控生产成本，实现原材料、能耗、人工等成本的精准核算和分析。',
    points: ['原材料成本核算', '能耗成本统计', '人工成本分析', '成本优化建议']
  },
  {
    title: '生产设备管理方案',
    icon: '🔧',
    desc: '设备全生命周期数字化管理，预防性维护、故障预警，保障生产线高效运转。',
    points: ['设备台账管理', '维护保养计划', '故障预警分析', '设备利用率统计']
  },
  {
    title: '生产效能管理方案',
    icon: '📊',
    desc: '全面监控分析生产效能指标，发现瓶颈、优化产能，提升整体生产效率。',
    points: ['OEE效率分析', '产能瓶颈识别', '生产节拍优化', '效能数据看板']
  },
  {
    title: '排产自动化解决方案',
    icon: '📅',
    desc: '智能排产系统根据订单、设备、人员等因素自动生成最优生产排程计划。',
    points: ['智能排产算法', '多约束条件优化', '排产计划自动调整', '产能负荷平衡']
  },
  {
    title: '生产质量管控方案',
    icon: '✅',
    desc: '全流程质量追溯和管控，从原材料到成品的完整质量管理体系。',
    points: ['质量检验管理', '质量追溯体系', '不合格品处理', '质量数据分析']
  },
  {
    title: '工人工资自动化核算方案',
    icon: '💳',
    desc: '根据生产数据自动核算工人工资，支持计件、计时等多种薪酬方式。',
    points: ['计件工资自动核算', '考勤数据整合', '多种薪酬模式', '工资明细报表']
  },
  {
    title: '无人值守解决方案',
    icon: '🔒',
    desc: '通过物联网、AI视觉等技术实现仓库和厂区的无人值守智能化管理。',
    points: ['智能门禁管控', 'AI视觉识别', '自动过磅称重', '异常自动报警']
  }
]
</script>
