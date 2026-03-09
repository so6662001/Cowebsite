import { createRouter, createWebHistory } from 'vue-router'

const routes = [
  {
    path: '/',
    name: 'Home',
    component: () => import('../views/Home.vue'),
    meta: {
      title: '南京普讯管理软件有限公司 - 钢铁贸易数字化解决方案提供商 | 钢贸宝ERP·库准MES·货袋子平台',
      description: '普讯软件深耕钢贸行业近20年，提供钢贸宝ERP、钢企通ERP、库准MES、智能财务等钢铁贸易数字化解决方案，服务近3000家钢贸企业，日活用户10000+。货袋子钢铁交易平台日均库存10万吨。',
      keywords: '钢贸宝ERP,钢铁贸易ERP,钢贸数字化,钢铁生产MES,库准仓库管理,货袋子钢铁交易平台,南京普讯,钢贸管理软件'
    }
  },
  {
    path: '/about',
    name: 'About',
    component: () => import('../views/About.vue'),
    meta: {
      title: '关于普讯 - 深耕钢铁贸易数字化近20年 | 南京普讯管理软件有限公司',
      description: '普讯深耕钢贸近20年，累计投入研发近5000万，员工近100人，服务团队覆盖全国各地。合作钢贸企业近3000家，荣获2家高新企业认证、5A级信用企业，拥有100多个产品著作权。',
      keywords: '普讯软件,钢贸数字化,钢铁行业ERP,关于普讯,公司介绍,发展历程'
    }
  },
  {
    path: '/products',
    name: 'Products',
    component: () => import('../views/Products.vue'),
    meta: {
      title: '产品中心 - 钢贸宝ERP·钢企通ERP·库准MES·智能财务·货袋子平台 | 普讯软件',
      description: '普讯软件产品矩阵：钢贸宝ERP助力钢贸企业经营数智化，钢企通ERP赋能钢铁生产企业，库准MES实现仓库智能管理，智能财务自动化核算，货袋子钢铁交易赋能平台日均库存10万吨。',
      keywords: '钢贸宝ERP,钢企通ERP,库准MES系统,智能财务系统,货袋子平台,钢铁行业软件'
    }
  },
  {
    path: '/solutions',
    name: 'Solutions',
    component: () => import('../views/Solutions.vue'),
    meta: {
      title: '解决方案 - 钢贸企业·生产制造企业数智化解决方案 | 普讯软件',
      description: '普讯提供钢贸企业和钢铁生产制造企业全套数智化解决方案，包括业务增长、降本增效、企业扩张、财务自动化、仓库管理、生产MES、AI数字员工等多种行业专属解决方案。',
      keywords: '钢贸解决方案,钢铁数字化转型,生产MES方案,仓库管理方案,降本增效,钢铁行业AI'
    }
  },
  {
    path: '/partners',
    name: 'Partners',
    component: () => import('../views/Partners.vue'),
    meta: {
      title: '合作企业 - 近3000家钢贸企业的共同选择 | 普讯软件',
      description: '普讯软件已服务近3000家钢贸企业、100+家库准合作企业、10+家生产MES合作企业。日活用户10000+人，货袋子日均线上库存10万吨。了解我们的战略合作客户。',
      keywords: '普讯合作企业,钢贸宝客户,钢铁行业合作,战略客户'
    }
  },
  {
    path: '/contact',
    name: 'Contact',
    component: () => import('../views/Contact.vue'),
    meta: {
      title: '联系我们 - 购前咨询·商务合作·售后服务 | 南京普讯管理软件有限公司',
      description: '联系南京普讯管理软件有限公司：总部地址上海，研发中心南京，销售中心南京。提供购前咨询、商务合作、售后服务支持，服务团队覆盖全国各地。',
      keywords: '联系普讯,普讯地址,钢贸宝咨询,商务合作,购前咨询'
    }
  },
  {
    path: '/trial',
    name: 'Trial',
    component: () => import('../views/Trial.vue'),
    meta: {
      title: '申请试用 - 免费体验钢贸宝ERP·库准MES·智能财务 | 普讯软件',
      description: '免费申请试用普讯软件产品：钢贸宝ERP、钢企通ERP、库准MES、智能财务系统。深耕钢贸行业近20年，以周为单位迭代升级，快速满足市场和客户需求。',
      keywords: '钢贸宝免费试用,ERP试用申请,钢铁行业软件试用,库准MES体验'
    }
  },
  {
    path: '/learning',
    name: 'Learning',
    component: () => import('../views/Learning.vue'),
    meta: {
      title: '学习中心 - 钢贸企业经营管理学习平台 | 普讯软件',
      description: '普讯学习中心提供钢贸企业经营管理培训、系统操作教程、行业知识分享。不仅是软件培训，更是一套经营企业的思路、理念以及落地方法。',
      keywords: '钢贸培训,ERP操作教程,企业经营管理学习,钢铁行业培训'
    }
  }
]

const router = createRouter({
  history: createWebHistory(),
  routes,
  scrollBehavior() {
    return { top: 0 }
  }
})

router.beforeEach((to, from, next) => {
  document.title = to.meta.title || '南京普讯管理软件有限公司'
  next()
})

export default router
